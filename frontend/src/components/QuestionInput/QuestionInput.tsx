import { useContext, useState } from 'react'
import { FontIcon, Stack, TextField } from '@fluentui/react'
import { SendRegular } from '@fluentui/react-icons'

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

export const QuestionInput = ({ onSend, disabled, placeholder, clearOnSend, conversationId }: Props) => {
  const [question, setQuestion] = useState<string>('')
  const [base64Image, setBase64Image] = useState<string | null>(null);
  const [documentText, setDocumentText] = useState<string | null>(null);
  const [documentName, setDocumentName] = useState<string | null>(null);
  const [documentImages, setDocumentImages] = useState<string[] | null>(null);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [isUploading, setIsUploading] = useState<boolean>(false);

  const appStateContext = useContext(AppStateContext)
  const OYD_ENABLED = appStateContext?.state.frontendSettings?.oyd_enabled || false;

  const handleImageUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setUploadError(null);

    if (file.type.startsWith('image/')) {
      await convertToBase64(file);
      return;
    }

    setIsUploading(true);
    try {
      const formData = new FormData();
      formData.append('file', file);
      const response = await fetch('/extract-document-text', {
        method: 'POST',
        body: formData
      });

      if (response.ok) {
        const data = await response.json();
        setDocumentText(data.text);
        setDocumentName(data.filename);
        setDocumentImages(data.images && data.images.length > 0 ? data.images : null);
      } else if (response.status === 413) {
        setUploadError('ファイルサイズが大きすぎます(上限30MB)。ファイルを圧縮するか、サイズを小さくしてから再度お試しください。');
      } else {
        let message = `アップロードに失敗しました(エラーコード: ${response.status})。`;
        try {
          const data = await response.json();
          if (data?.error) message = data.error;
        } catch {
          // レスポンスがJSONでない場合はそのまま既定のメッセージを使う
        }
        setUploadError(message);
      }
    } catch (error) {
      console.error('Error:', error);
      setUploadError('ファイルのアップロード中にエラーが発生しました。ネットワーク状況を確認し、再度お試しください。');
    } finally {
      setIsUploading(false);
      event.target.value = '';
    }
  };

  const convertToBase64 = async (file: Blob) => {
    try {
      const resizedBase64 = await resizeImage(file, 800, 800);
      setBase64Image(resizedBase64);
    } catch (error) {
      console.error('Error:', error);
    }
  };

  const sendQuestion = () => {
    if (disabled || !question.trim()) {
      return
    }

    const documentContext = documentText
      ? `以下はアップロードされたファイル「${documentName}」の内容です。\n\n${documentText}\n\n---\n\n上記の内容を踏まえて、次の質問に答えてください。\n\n質問: ${question}`
      : undefined;

    // ファイルの中身は表示・履歴に残さないが、添付した事実だけは分かるようにファイル名を付記する
    const displayedQuestion = documentName ? `${question}\n\n📎 ${documentName}` : question;

    const questionTest: ChatMessage["content"] = base64Image ? [{ type: "text", text: question }, { type: "image_url", image_url: { url: base64Image } }] : displayedQuestion.toString();

    if (conversationId && questionTest !== undefined) {
      onSend(questionTest, conversationId, documentContext, documentImages ?? undefined)
      setBase64Image(null)
    } else {
      onSend(questionTest, undefined, documentContext, documentImages ?? undefined)
      setBase64Image(null)
    }
    setDocumentText(null)
    setDocumentName(null)
    setDocumentImages(null)
    setUploadError(null)

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

  return (
    <Stack horizontal className={styles.questionInputContainer}>
      <TextField
        className={styles.questionInputTextArea}
        placeholder={placeholder}
        multiline
        resizable={false}
        borderless
        value={question}
        onChange={onQuestionChange}
        onKeyDown={onEnterPress}
      />
      {!OYD_ENABLED && (
        <div className={styles.fileInputContainer}>
          <input
            type="file"
            id="fileInput"
            onChange={(event) => handleImageUpload(event)}
            accept="image/*,.txt,.md,.json,.html,.htm,.pdf,.docx,.xlsx,.xls,.pptx"
            className={styles.fileInput}
          />
          <label htmlFor="fileInput" className={styles.fileLabel} aria-label='Upload Image'>
            <FontIcon
              className={styles.fileIcon}
              iconName={'PhotoCollection'}
              aria-label='Upload Image'
            />
          </label>
        </div>)}
      {base64Image && <img className={styles.uploadedImage} src={base64Image} alt="Uploaded Preview" />}
      {isUploading && (
        <div style={{ fontSize: '12px', alignSelf: 'center', marginRight: '8px' }}>
          アップロード中...
        </div>
      )}
      {documentName && !base64Image && !isUploading && (
        <div aria-label={`Attached file: ${documentName}`} style={{ fontSize: '12px', alignSelf: 'center', marginRight: '8px' }}>
          📎 {documentName}
        </div>
      )}
      {uploadError && (
        <div role="alert" style={{ fontSize: '12px', alignSelf: 'center', marginRight: '8px', color: '#a4262c', maxWidth: '240px' }}>
          {uploadError}
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
