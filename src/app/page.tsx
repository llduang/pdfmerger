import { MergeTool } from '@/components/merge-tool';

export default function Home() {
  return (
    <main className="min-h-screen bg-gradient-to-br from-violet-600 via-purple-600 to-purple-700 py-6 md:py-10 px-4 md:px-6">
      {/* Header */}
      <header className="text-center mb-6 md:mb-8">
        <div className="flex items-center justify-center gap-3 mb-2">
          <h1 className="text-2xl md:text-4xl font-bold text-white tracking-tight drop-shadow-lg">
            PDF、图片与 Word 合并工具
          </h1>
        </div>
        <p className="text-sm md:text-base text-white/80 max-w-2xl mx-auto">
          上传 PDF、图片和 Word 文档，智能合并为一个可打印的 PDF 文件
        </p>
      </header>

      {/* Main Card */}
      <div className="bg-white rounded-2xl shadow-2xl p-5 md:p-8 max-w-4xl mx-auto">
        <MergeTool />
      </div>

      {/* Footer */}
      <footer className="text-center mt-6 text-white/50 text-xs">
        <p>所有文件处理均在本地浏览器完成，不会上传到任何服务器</p>
      </footer>
    </main>
  );
}
