# Paiban（排版）

中文长文档转 WPS 原生可编辑 PPT 的版式引擎。

Paiban 把 markdown / docx / txt 长文档转成 `.pptx`，并在 OOXML 层应用一套确定性的中文版式规则：字体别名回退（微软雅黑 / 宋体 / 黑体 / 思源黑体落到 `<a:latin>` `<a:ea>` `<a:cs>` 三个字体槽）、标点挤压与避头尾、首行缩进 2 字符、行距 1.5。产出的母版与占位符是真实形状，在 WPS 里双击即可继续改字加页，保存后版式不崩。

CLI、模板族与完整文档仍在实现中；本文件是仓库骨架版本，完整的中文 README 与 `README.en.md` 随后补齐。

## License

[MIT](./LICENSE) © 2026 SuperMarioYL
