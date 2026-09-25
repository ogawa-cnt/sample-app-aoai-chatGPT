import { useState } from 'react'
import { DefaultButton } from '@fluentui/react'

interface Props {
  format: string
  filename: string
  content: string
}

export const GeneratedFileButton = ({ format, filename, content }: Props) => {
  const [isDownloading, setIsDownloading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const handleDownload = async () => {
    setIsDownloading(true)
    setError(null)
    try {
      const response = await fetch('/generate-document', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ format, filename, content })
      })

      if (!response.ok) {
        let message = `生成に失敗しました(エラーコード: ${response.status})。`
        try {
          const data = await response.json()
          if (data?.error) message = data.error
        } catch {
          // レスポンスがJSONでない場合はそのまま既定のメッセージを使う
        }
        throw new Error(message)
      }

      const data = await response.json()
      const binary = atob(data.data)
      const bytes = new Uint8Array(binary.length)
      for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i)
      }
      const blob = new Blob([bytes])
      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      link.download = data.filename || filename
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
      URL.revokeObjectURL(url)
    } catch (e) {
      console.error('Error:', e)
      setError(e instanceof Error ? e.message : 'ダウンロードに失敗しました。')
    } finally {
      setIsDownloading(false)
    }
  }

  return (
    <div style={{ margin: '8px 0' }}>
      <DefaultButton onClick={handleDownload} disabled={isDownloading}>
        {isDownloading ? '生成中...' : `📥 ${filename} をダウンロード`}
      </DefaultButton>
      {error && <div style={{ color: '#a4262c', fontSize: '12px', marginTop: '4px' }}>{error}</div>}
    </div>
  )
}
