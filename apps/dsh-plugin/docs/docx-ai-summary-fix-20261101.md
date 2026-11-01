# Word 总结读取修复

2026-11-01。文档工作台右侧 Chat 发送“总结当前页内容”或“总结当前文档内容”，Word 正文已显示仍提示编辑器未加载。

根因：`easyview_read_document` 仅支持 Markdown、TXT 和 XLSX；DOCX 没有正文读取适配，且在判断格式前错误套用了文本编辑器的 generation 检查。补齐 generation 后仍会得到格式不支持错误。

修复：Host 从已授权的 DOCX 原件或当前工作台的已确认恢复草稿提取正文，保留段落和表格文字，解析 XML 实体，跳过删除修订和字段指令。读取仍核对文件身份、原件哈希、草稿基线及已提供的 generation。响应标记来源、版本、整篇文档范围和截断状态。无需为读取原件等待文本编辑器；已挂载 Word 编辑器发送前会等待恢复和草稿确认。

真实界面验证同时复现了刚输入后立即总结会报 `editor_revision_changed`：后台序列化遇到新输入时会重试，但显式 flush 不会。共享 checkpoint queue 现在对这项捕获失效重新读取最新版本；真实存储失败仍向上传递，丢失 ACK 时仍重用同一快照。当前环境 `node_modules/@easyview/preview-ui` 指向 `C:/workspace/dsh/EasyView_Md-dsh/packages/preview-ui`，因此共享仓库及插件仓库内副本均包含相同修复和测试。

验证：

- `DSH_TEST_RUNTIME=C:/workspace/dsh/deepseek-harness node --import tsx/esm --test tests/*.test.mjs`：44 项通过，覆盖 Word 挂载/未挂载与原件/草稿的四种组合、两个原始总结请求、表格、XML 实体、删除修订及原件未被写入。
- `node node_modules/vitest/vitest.mjs run packages/preview-ui/src/editorCheckpointQueue.test.ts packages/preview-ui/src/editorLifecycle.test.ts`：8 项通过，包含 flush 自己启动捕获及等待后台捕获时的新输入竞争。
- Host/Frame 两项 TypeScript 检查及构建通过。
- `scripts/probe-ai-controlled.mjs`：真实 Word 编辑器和聊天栏发送上述两个请求；第二次发送前输入未保存文字，确认工具读取最新草稿且原件字节不变。原 Markdown 编辑/撤销/恢复与 XLSX 提案/拒绝/应用/撤销也通过。模型响应受控，不代表真实模型输出质量验收。结果见 `docs/evidence/e05-controlled-ai.json`，安装资源验证见 `docs/evidence/e05-package-controlled-ai.json`。

“当前页”在本修复中指当前打开的文档标签。工具返回整篇正文，不提供 Word 排版后的物理页码，模型工具描述已明确该范围，禁止声称只总结某个物理页。正文上下文最多 256,000 字符，超限明确标记截断；正文 XML 解压上限为 16 MiB。

本机增量使用 `scripts/install-docx-ai-update.mjs`，要求显式 `DSH_DESKTOP_PROFILE`，先核对静态资源兼容性，再备份并更新 Host、preset、renderer 及所需 JS 块和资源清单。保留旧块以供正在运行的旧页面使用；异常回滚已写文件。安装位置、备份及散列见 `docs/evidence/docx-ai-desktop-install.json`。不修改用户配置或文档，不主动重启，重启桌面应用后加载修复。
