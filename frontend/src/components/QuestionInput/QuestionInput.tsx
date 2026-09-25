import { useContext, useState } from 'react'
import { Stack, TextField } from '@fluentui/react'
import { DocumentAddRegular, SendRegular } from '@fluentui/react-icons'
import uuid from 'react-uuid'

import Send from '../../assets/Send.svg'

import styles from './QuestionInput.module.css'
import { ChatMessage } from '../../api'
import { AppStateContext } from '../../state/AppProvider'
import { resizeImage } from '../../utils/resizeImage'

interface Props {
  onSend: (question: ChatMessage['content'], id?: string, documentContext?: string, documentImages?: string[]) => void
  disabled: boolean
  placeholder?: string
  clearOnSend?: boolean
  conversationId?: string
}

interface Attachment {
  id: string
  name: string
  kind: 'image' | 'document'
  status: 'uploading' | 'ready' | 'error'
  errorMessage?: string
  imageDataUrl?: string
  documentText?: string
  documentImages?: string[]
}

const MAX_ATTACHMENTS = 5
const MAX_COMBINED_CHARS = 200000
const MAX_COMBINED_IMAGES = 10

export const QuestionInput = ({ onSend, disabled, placeholder, clearOnSend, conversationId }: Props) => {
  const [question, setQuestion] = useState<string>('')
  const [attachments, setAttachments] = useState<Attachment[]>([])
  const [globalError, setGlobalError] = useState<string | null>(null)
  const [isDraggingOver, setIsDraggingOver] = useState<boolean>(false)

  const appStateContext = useContext(AppStateContext)
  const OYD_ENABLED = appStateContext?.state.frontendSettings?.oyd_enabled || false;

  const processFile = async (file: File, id: string) => {
    if (file.type.startsWith('image/')) {
      try {
        const resizedBase64 = await resizeImage(file, 800, 800)
        setAttachments(prev => prev.map(a => (a.id === id ? { ...a, status: 'ready', imageDataUrl: resizedBase64 } : a)))
      } catch (error) {
        console.error('Error:', error)
        setAttachments(prev =>
          prev.map(a => (a.id === id ? { ...a, status: 'error', errorMessage: '画像の読み込みに失敗しました。' } : a))
        )
      }
      return
    }

    try {
      const formData = new FormData()
      formData.append('file', file)
      const response = await fetch('/extract-document-text', {
        method: 'POST',
        body: formData
      })

      if (response.ok) {
        const data = await response.json()
        setAttachments(prev =>
          prev.map(a =>
            a.id === id
              ? { ...a, status: 'ready', name: data.filename || a.name, documentText: data.text, documentImages: data.images || [] }
              : a
          )
        )
      } else if (response.status === 413) {
        setAttachments(prev =>
          prev.map(a => (a.id === id ? { ...a, status: 'error', errorMessage: 'ファイルサイズが大きすぎます(上限30MB)。' } : a))
        )
      } else {
        let message = `アップロードに失敗しました(エラーコード: ${response.status})。`
        try {
          const data = await response.json()
          if (data?.error) message = data.error
        } catch {
          // レスポンスがJSONでない場合はそのまま既定のメッセージを使う
        }
        setAttachments(prev => prev.map(a => (a.id === id ? { ...a, status: 'error', errorMessage: message } : a)))
      }
    } catch (error) {
      console.error('Error:', error)
      setAttachments(prev =>
        prev.map(a =>
          a.id === id ? { ...a, status: 'error', errorMessage: 'ファイルのアップロード中にエラーが発生しました。' } : a
        )
      )
    }
  }

  const addFiles = (files: FileList | File[]) => {
    if (disabled) return

    const fileArray = Array.from(files)
    if (fileArray.length === 0) return

    setGlobalError(null)

    const availableSlots = MAX_ATTACHMENTS - attachments.length
    if (availableSlots <= 0) {
      setGlobalError(`最大${MAX_ATTACHMENTS}ファイルまで添付できます。`)
      return
    }

    const filesToProcess = fileArray.slice(0, availableSlots)
    if (fileArray.length > filesToProcess.length) {
      setGlobalError(`最大${MAX_ATTACHMENTS}ファイルまで添付できるため、一部のファイルは追加されませんでした。`)
    }

    const newAttachments: Attachment[] = filesToProcess.map(file => ({
      id: uuid(),
      name: file.name || '画像',
      kind: file.type.startsWith('image/') ? 'image' : 'document',
      status: 'uploading'
    }))

    setAttachments(prev => [...prev, ...newAttachments])
    filesToProcess.forEach((file, index) => processFile(file, newAttachments[index].id))
  }

  const removeAttachment = (id: string) => {
    setAttachments(prev => prev.filter(a => a.id !== id))
    setGlobalError(null)
  }

  const handleFileInputChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files && event.target.files.length > 0) {
      addFiles(event.target.files)
    }
    event.target.value = ''
  }

  const handleDragOver = (event: React.DragEvent) => {
    event.preventDefault()
    if (!disabled) setIsDraggingOver(true)
  }

  const handleDragLeave = (event: React.DragEvent) => {
    event.preventDefault()
    setIsDraggingOver(false)
  }

  const handleDrop = (event: React.DragEvent) => {
    event.preventDefault()
    setIsDraggingOver(false)
    if (event.dataTransfer.files && event.dataTransfer.files.length > 0) {
      addFiles(event.dataTransfer.files)
    }
  }

  const handlePaste = (event: React.ClipboardEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const items = event.clipboardData?.items
    if (!items) return

    const imageFiles: File[] = []
    for (let i = 0; i < items.length; i++) {
      const item = items[i]
      if (item.type.startsWith('image/')) {
        const file = item.getAsFile()
        if (file) imageFiles.push(file)
      }
    }

    if (imageFiles.length > 0) {
      event.preventDefault()
      addFiles(imageFiles)
    }
    // 画像以外の通常のテキスト貼り付けは、そのままデフォルトの動作に任せる
  }

  const sendQuestion = () => {
    if (disabled || !question.trim()) {
      return
    }

    const readyAttachments = attachments.filter(a => a.status === 'ready')

    const documentSections = readyAttachments
      .filter(a => a.kind === 'document' && a.documentText)
      .map(a => `## 添付ファイル: ${a.name}\n\n${a.documentText}`)

    let combinedDocumentText = documentSections.join('\n\n---\n\n')
    if (combinedDocumentText.length > MAX_COMBINED_CHARS) {
      combinedDocumentText =
        combinedDocumentText.slice(0, MAX_COMBINED_CHARS) + '\n\n...(以下省略、文字数上限のため切り捨てられました)'
    }

    const combinedImages = readyAttachments
      .flatMap(a => (a.kind === 'image' && a.imageDataUrl ? [a.imageDataUrl] : a.documentImages || []))
      .slice(0, MAX_COMBINED_IMAGES)

    const hasAttachmentContent = combinedDocumentText.length > 0 || combinedImages.length > 0

    const documentContext = hasAttachmentContent
      ? `以下は添付されたファイルの内容です。\n\n${combinedDocumentText}\n\n---\n\n上記の内容を踏まえて、次の質問に答えてください。\n\n質問: ${question}`
      : undefined

    // ファイルの中身は表示・履歴に残さないが、添付した事実だけは分かるようにファイル名を付記する
    const attachmentSuffix =
      readyAttachments.length > 0 ? '\n\n' + readyAttachments.map(a => `📌 ${a.name}`).join(' ') : ''
    const displayedQuestion = `${question}${attachmentSuffix}`

    onSend(displayedQuestion, conversationId, documentContext, combinedImages.length > 0 ? combinedImages : undefined)

    // 送信に使われた(準備完了の)添付だけを消す。失敗したものは残し、黙って消えないようにする
    setAttachments(prev => prev.filter(a => a.status !== 'ready'))
    setGlobalError(null)

    if (clearOnSend) {
      setQuestion('')
    }
  }

  const onEnterPress = (ev: React.KeyboardEvent<Element>) => {
    if (ev.key === 'Enter' && !ev.shiftKey && !(ev.nativeEvent?.isComposing === true)) {
      ev.preventDefault()
      sendQuestion()
    }
  }

  const onQuestionChange = (_ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    setQuestion(newValue || '')
  }

  const sendQuestionDisabled = disabled || !question.trim()

  const attachmentStatusIcon = (attachment: Attachment) => {
    if (attachment.status === 'uploading') return '⏳'
    if (attachment.status === 'error') return '⚠️'
    return '📌'
  }

  return (
    <Stack
      horizontal
      className={styles.questionInputContainer}
      style={attachments.length > 0 ? { height: '168px' } : undefined}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}>
      <Stack className={styles.questionInputMain}>
        {attachments.length > 0 && (
          <div className={styles.attachmentsRow}>
            {attachments.map(attachment => (
              <div
                key={attachment.id}
                className={styles.attachmentChip}
                title={attachment.status === 'error' ? attachment.errorMessage : attachment.name}>
                {attachment.kind === 'image' && attachment.imageDataUrl ? (
                  <img src={attachment.imageDataUrl} className={styles.attachmentThumb} alt={attachment.name} />
                ) : (
                  <span>{attachmentStatusIcon(attachment)}</span>
                )}
                <span className={styles.attachmentName}>{attachment.name}</span>
                <button
                  type="button"
                  className={styles.attachmentRemove}
                  onClick={() => removeAttachment(attachment.id)}
                  aria-label={`Remove ${attachment.name}`}>
                  ×
                </button>
              </div>
            ))}
          </div>
        )}
        <TextField
          className={styles.questionInputTextArea}
          placeholder={placeholder}
          multiline
          resizable={false}
          borderless
          value={question}
          onChange={onQuestionChange}
          onKeyDown={onEnterPress}
          onPaste={handlePaste}
        />
      </Stack>
      {isDraggingOver && (
        <div className={styles.dragOverlay}>
          <span>ここにファイルをドロップして添付</span>
        </div>
      )}
      {!OYD_ENABLED && (
        <div className={styles.fileInputContainer}>
          <input
            type="file"
            id="fileInput"
            multiple
            onChange={handleFileInputChange}
            accept="image/*,.txt,.md,.json,.html,.htm,.pdf,.docx,.xlsx,.xls,.pptx"
            className={styles.fileInput}
          />
          <label htmlFor="fileInput" className={styles.fileLabel} aria-label='Attach file'>
            <DocumentAddRegular className={styles.fileIcon} aria-label='Attach file' />
          </label>
        </div>)}
      {globalError && (
        <div role="alert" className={styles.globalError}>
          {globalError}
        </div>
      )}
      <div
        className={styles.questionInputSendButtonContainer}
        role="button"
        tabIndex={0}
        aria-label="Ask question button"
        onClick={sendQuestion}
        onKeyDown={e => (e.key === 'Enter' || e.key === ' ' ? sendQuestion() : null)}>
        {sendQuestionDisabled ? (
          <SendRegular className={styles.questionInputSendButtonDisabled} />
        ) : (
          <img src={Send} className={styles.questionInputSendButton} alt="Send Button" />
        )}
      </div>
      <div className={styles.questionInputBottomBorder} />
    </Stack>
  )
}
