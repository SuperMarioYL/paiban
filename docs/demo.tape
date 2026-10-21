# docs/demo.tape — 30 秒终端实录（vhs 脚本）
#
# 录的就是 README 快速开始那条路：长文档 → .pptx → 版式 lint → 属性确实落在 OOXML 里。
# 在仓库根目录运行：vhs docs/demo.tape（需要 vhs + ttyd + ffmpeg）
#
# 只录真实终端路径。WPS 往返（打开 / 改字 / 加页 / 保存）与和别的工具的并排对比需要本机
# 装有 WPS 和那个工具，不在录制范围内——见 docs/demo-results.json 的 limitations。

Output "assets/demo.gif"

Set FontFamily "Menlo, Noto Sans Mono CJK SC, PingFang SC, Hiragino Sans GB"
Set FontSize 15
Set Width 1400
Set Height 660
Set Padding 20
Set Margin 8
Set MarginFill "#1e1e2e"
Set BorderRadius 6
Set Theme "Catppuccin Mocha"
Set TypingSpeed 35ms
Set Framerate 20
Set CursorBlink false

Hide
Type `clear`
Enter
Show

Type `paiban build fixtures/sample-longdoc.md -o 课件.pptx --template courseware`
Enter
Sleep 2s

Type `paiban verify 课件.pptx`
Enter
Sleep 7s

Type `unzip -p 课件.pptx ppt/slides/slide2.xml | grep -oE '(a:ea typeface|spcPct val|indent)="[^"]*"' | sort -u`
Enter
Sleep 5s

Sleep 2s
