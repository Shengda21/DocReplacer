# 如何使用 GitHub 自动生成 Mac 和 Ubuntu 软件

由于您是电脑新手，我为您准备了“保姆级”的使用教程。您只需要按照以下步骤，把代码传到 GitHub，GitHub 就会免费帮您打包出 Windows、Mac 和 Ubuntu 三个版本的软件。

## 第 1 步：安装 Git（如果您还没安装）
1. 去官网下载 Git：https://git-scm.com/downloads
2. 一路点击“下一步 (Next)”默认安装即可。

## 第 2 步：在 GitHub 上创建一个新仓库
1. 登录您的 GitHub 账号 (https://github.com/)。
2. 点击右上角的 **+** 号，选择 **New repository**。
3. **Repository name** 随便填一个名字，比如 `DocReplacer`。
4. **Public** 或 **Private** 都可以（Private 别人看不到，但每个月免费打包时长有限额，一般自己用选 Private 即可）。
5. 不要勾选其他东西，直接点击绿色的 **Create repository** 按钮。
6. 创建成功后，您会看到一个类似于 `https://github.com/您的用户名/DocReplacer.git` 的链接，请复制备用。

## 第 3 步：把电脑上的代码推送到 GitHub
1. 打开 `c:\code\替换` 文件夹。
2. 在文件夹空白处 **右键**，如果您使用的是 Windows 11，点击“显示更多选项”，然后选择 **Open Git Bash here**（或在文件夹地址栏输入 `cmd` 按回车）。
3. 在弹出的黑框框里，依次输入以下命令，**每输完一行按一次回车**：

```shell
git init
git add .
git commit -m "第一次提交代码"
git branch -M main
git remote add origin 您刚刚在第2步复制的链接
git push -u origin main
```
> 第二行 `git add .` 后面有一个点，不要漏掉。

## 第 4 步：等待 GitHub 自动打包
1. 回到 GitHub 网页，点击上面菜单的 **Actions** 标签夹。
2. 您会看到一个叫 `Build Multi-Platform Executables` 的任务正在转圈圈运行。
3. 等待大约 2-5 分钟，等它变成绿色的勾 ✅，说明打包完成了。

## 第 5 步：下载软件
1. 点击那个变成了绿色勾的任务。
2. 往下滚动网页，您会看到 **Artifacts** 区域（这就是打包好的成品）。
3. 里面会列出：
   - `DocumentReplacer_macOS` （Mac版）
   - `DocumentReplacer_Ubuntu` （Ubuntu Linux版）
   - `DocumentReplacer_Windows.exe` （Windows版）
4. 直接点击对应平台的名字，就可以下载压缩包了。解压后即可在对应的电脑上双击运行。

---
如果您在操作过程中遇到任何报错或看不懂的地方，随时把屏幕上的提示发给我！
