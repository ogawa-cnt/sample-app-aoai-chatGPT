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

export function parseAnswer(answer: AskResponse, isStreaming: boolean = false): ParsedAnswer {
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
  // まだ閉じタグ([[/FILE]])が届いていないブロックが残っている場合、
  // 応答がまだストリーミング中なら「生成中」、応答が終わっているのに閉じていない場合は
  // 出力が途中で切れて失敗したとみなし、その旨を表示する(無限に「生成中」のままにしない)
  const danglingIndex = answerText.indexOf('[[FILE:')
  if (danglingIndex !== -1) {
    const notice = isStreaming
      ? '_ファイルを生成中..._'
      : '_⚠️ ファイルの生成が完了しませんでした(内容が長すぎた可能性があります)。お手数ですが、対象を絞って再度お試しください。_'
    answerText = answerText.slice(0, danglingIndex) + '\n\n' + notice
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
