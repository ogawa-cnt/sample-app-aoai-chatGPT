import { cloneDeep } from 'lodash'

import { AskResponse, Citation } from '../../api'

export type GeneratedFile = {
  format: string
  filename: string
  content: string
}

export type ParsedAnswer = {
  citations: Citation[]
  markdownFormatText: string,
  generated_chart: string | null,
  generatedFiles: GeneratedFile[]
} | null

export const enumerateCitations = (citations: Citation[]) => {
  const filepathMap = new Map()
  for (const citation of citations) {
    const { filepath } = citation
    let part_i = 1
    if (filepathMap.has(filepath)) {
      part_i = filepathMap.get(filepath) + 1
    }
    filepathMap.set(filepath, part_i)
    citation.part_index = part_i
  }
  return citations
}

export function parseAnswer(answer: AskResponse): ParsedAnswer {
  if (typeof answer.answer !== "string") return null
  let answerText = answer.answer

  // ダウンロード可能なファイルの生成ブロック([[FILE:拡張子:ファイル名]]...[[/FILE]])を抜き出し、
  // 本文には残さずダウンロードボタンとして表示する
  const generatedFiles: GeneratedFile[] = []
  const fileBlockPattern = /\[\[FILE:([a-zA-Z0-9]+):([^\]\n]+)\]\]([\s\S]*?)\[\[\/FILE\]\]/g
  answerText = answerText.replace(fileBlockPattern, (_match, format, filename, content) => {
    generatedFiles.push({ format: format.trim().toLowerCase(), filename: filename.trim(), content: content.trim() })
    return ''
  })
  // ストリーミング中で、まだ閉じタグが届いていない生成ブロックは、そのまま表示せず一時的な案内に置き換える
  const danglingIndex = answerText.indexOf('[[FILE:')
  if (danglingIndex !== -1) {
    answerText = answerText.slice(0, danglingIndex) + '\n\n_ファイルを生成中..._'
  }

  const citationLinks = answerText.match(/\[(doc\d\d?\d?)]/g)

  const lengthDocN = '[doc'.length

  let filteredCitations = [] as Citation[]
  let citationReindex = 0
  citationLinks?.forEach(link => {
    // Replacing the links/citations with number
    const citationIndex = link.slice(lengthDocN, link.length - 1)
    const citation = cloneDeep(answer.citations[Number(citationIndex) - 1]) as Citation
    if (!filteredCitations.find(c => c.id === citationIndex) && citation) {
      answerText = answerText.replaceAll(link, ` ^${++citationReindex}^ `)
      citation.id = citationIndex // original doc index to de-dupe
      citation.reindex_id = citationReindex.toString() // reindex from 1 for display
      filteredCitations.push(citation)
    }
  })

  filteredCitations = enumerateCitations(filteredCitations)

  return {
    citations: filteredCitations,
    markdownFormatText: answerText,
    generated_chart: answer.generated_chart,
    generatedFiles
  }
}
