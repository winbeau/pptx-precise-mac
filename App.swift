import Cocoa
import UniformTypeIdentifiers

final class AppDelegate: NSObject, NSApplicationDelegate, NSWindowDelegate {
    private var window: NSWindow!
    private let filesLabel = NSTextField(wrappingLabelWithString: "选择 PPTX，或将文件、文件夹拖到应用图标上。")
    private let folderLabel = NSTextField(wrappingLabelWithString: "尚未选择保存位置")
    private let statusLabel = NSTextField(labelWithString: "准备就绪")
    private let log = NSTextView()
    private let spinner = NSProgressIndicator()
    private var files: [URL] = []
    private var output: URL?
    private var busy = false
    private var controls: [NSButton] = []

    func applicationDidFinishLaunching(_ notification: Notification) {
        NSApp.setActivationPolicy(.regular)
        let menu = NSMenu()
        let appItem = NSMenuItem(); menu.addItem(appItem)
        let submenu = NSMenu(); appItem.submenu = submenu
        submenu.addItem(withTitle: "退出 PPTX Precise", action: #selector(NSApplication.terminate(_:)), keyEquivalent: "q")
        NSApp.mainMenu = menu
        window = NSWindow(contentRect: NSRect(x: 0, y: 0, width: 720, height: 510),
                          styleMask: [.titled, .closable, .miniaturizable], backing: .buffered, defer: false)
        window.title = "PPTX Precise"
        window.delegate = self
        window.isReleasedWhenClosed = false
        let view = window.contentView!
        let title = NSTextField(labelWithString: "PPTX Precise")
        title.font = .systemFont(ofSize: 27, weight: .semibold)
        title.frame = NSRect(x: 28, y: 449, width: 420, height: 38); view.addSubview(title)
        let subtitle = NSTextField(labelWithString: "PowerPoint 原生排版 · 原图恢复与校验 · 全部在本机完成")
        subtitle.textColor = .secondaryLabelColor
        subtitle.frame = NSRect(x: 28, y: 422, width: 650, height: 22); view.addSubview(subtitle)
        let select = NSButton(title: "选择 PPTX…", target: self, action: #selector(selectFiles))
        select.frame = NSRect(x: 24, y: 370, width: 132, height: 32); view.addSubview(select)
        let clear = NSButton(title: "清空", target: self, action: #selector(clearFiles))
        clear.frame = NSRect(x: 158, y: 370, width: 72, height: 32); view.addSubview(clear)
        filesLabel.frame = NSRect(x: 28, y: 316, width: 660, height: 50); view.addSubview(filesLabel)
        let folder = NSButton(title: "保存位置…", target: self, action: #selector(selectOutput))
        folder.frame = NSRect(x: 24, y: 274, width: 132, height: 32); view.addSubview(folder)
        folderLabel.font = .systemFont(ofSize: 12)
        folderLabel.frame = NSRect(x: 166, y: 272, width: 520, height: 34); view.addSubview(folderLabel)
        let start = NSButton(title: "转换为 PDF", target: self, action: #selector(convert))
        start.bezelStyle = .rounded
        start.keyEquivalent = "\r"
        start.frame = NSRect(x: 548, y: 222, width: 144, height: 34); view.addSubview(start)
        spinner.style = .spinning; spinner.isDisplayedWhenStopped = false
        spinner.frame = NSRect(x: 28, y: 229, width: 18, height: 18); view.addSubview(spinner)
        statusLabel.frame = NSRect(x: 54, y: 226, width: 480, height: 22); view.addSubview(statusLabel)
        let scroll = NSScrollView(frame: NSRect(x: 28, y: 52, width: 660, height: 158))
        scroll.hasVerticalScroller = true; scroll.borderType = .bezelBorder
        log.isEditable = false; log.isSelectable = true
        log.font = .monospacedSystemFont(ofSize: 11, weight: .regular)
        log.autoresizingMask = [.width]; log.textContainerInset = NSSize(width: 9, height: 9)
        log.frame = NSRect(x: 0, y: 0, width: 640, height: 158)
        log.isVerticallyResizable = true
        log.textContainer?.widthTracksTextView = true
        scroll.documentView = log; view.addSubview(scroll)
        let footer = NSTextField(labelWithString: "需安装 PowerPoint。已有同名输出会自动备份，源稿的未保存修改需先保存。")
        footer.font = .systemFont(ofSize: 11); footer.textColor = .secondaryLabelColor
        footer.frame = NSRect(x: 28, y: 19, width: 660, height: 20); view.addSubview(footer)
        controls = [select, clear, folder, start]
        refresh()
        window.center(); window.makeKeyAndOrderFront(nil)
        NSApp.activate(ignoringOtherApps: true)
    }

    private func refresh() {
        guard window != nil else { return }
        if files.isEmpty { filesLabel.stringValue = "选择 PPTX，或将文件、文件夹拖到应用图标上。" }
        else { filesLabel.stringValue = "已选 \(files.count) 项：" + files.map { $0.lastPathComponent }.joined(separator: "、") }
        folderLabel.stringValue = output?.path ?? "尚未选择保存位置"
        controls.forEach { $0.isEnabled = !busy }
        controls.last?.isEnabled = !busy && !files.isEmpty && output != nil
    }

    private func addFiles(_ urls: [URL]) {
        guard !busy else { return }
        for url in urls {
            if !files.contains(url) { files.append(url) }
        }
        if output == nil, let first = files.first { output = first.deletingLastPathComponent() }
        refresh()
    }

    @objc private func selectFiles() {
        let panel = NSOpenPanel()
        panel.allowsMultipleSelection = true
        panel.allowedContentTypes = [UTType(filenameExtension: "pptx")!]
        panel.canChooseDirectories = false
        panel.beginSheetModal(for: window) { response in
            if response == .OK { self.addFiles(panel.urls) }
        }
    }

    @objc private func clearFiles() { files = []; refresh() }

    @objc private func selectOutput() {
        let panel = NSOpenPanel()
        panel.canChooseFiles = false; panel.canChooseDirectories = true
        panel.canCreateDirectories = true
        panel.prompt = "选择保存位置"
        panel.directoryURL = output
        panel.beginSheetModal(for: window) { response in
            if response == .OK { self.output = panel.url; self.refresh() }
        }
    }

    @objc private func convert() {
        guard !busy, !files.isEmpty, let output = output,
              let resource = Bundle.main.resourceURL else { return }
        busy = true; refresh(); spinner.startAnimation(nil)
        statusLabel.stringValue = "正在导出和校验…"
        log.string = "首次运行如出现 PowerPoint 文件访问提示，请选择工具的 Work 文件夹并允许访问。\n"
        let process = Process()
        process.executableURL = resource.appendingPathComponent("bin/precise-pptx/precise-pptx")
        process.arguments = files.map { $0.path } + ["--output", output.path, "--copy-pptx", "--overwrite"]
        let pipe = Pipe(); process.standardOutput = pipe; process.standardError = pipe
        do { try process.run() }
        catch {
            busy = false; spinner.stopAnimation(nil); refresh()
            statusLabel.stringValue = "无法启动转换"
            log.string = error.localizedDescription
            return
        }
        DispatchQueue.global(qos: .userInitiated).async {
            let data = pipe.fileHandleForReading.readDataToEndOfFile()
            process.waitUntilExit()
            let text = String(data: data, encoding: .utf8) ?? "无法读取转换日志"
            DispatchQueue.main.async {
                self.busy = false; self.spinner.stopAnimation(nil); self.refresh()
                self.statusLabel.stringValue = process.terminationStatus == 0 ? "转换完成" : "转换未完成，请查看下方提示"
                self.log.string = text
                self.window.makeKeyAndOrderFront(nil)
            }
        }
    }

    func application(_ sender: NSApplication, openFiles filenames: [String]) {
        addFiles(filenames.map { URL(fileURLWithPath: $0) })
        window?.makeKeyAndOrderFront(nil)
        sender.reply(toOpenOrPrint: .success)
    }
    func applicationShouldTerminateAfterLastWindowClosed(_ sender: NSApplication) -> Bool { !busy }
    func applicationShouldTerminate(_ sender: NSApplication) -> NSApplication.TerminateReply { busy ? .terminateCancel : .terminateNow }
    func windowShouldClose(_ sender: NSWindow) -> Bool { !busy }
}
let app = NSApplication.shared
let delegate = AppDelegate()
app.delegate = delegate
app.run()
