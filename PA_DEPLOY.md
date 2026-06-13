# PythonAnywhere 部署指南（免费，无需绑卡）

## 1. 注册
访问 https://www.pythonanywhere.com/registration/register/beginner/
创建免费账号（记下用户名）

## 2. 克隆代码
登录后打开 Consoles → Bash，输入：

```bash
git clone https://github.com/Bryce-queen/temu-audit.git
```

## 3. 安装依赖
```bash
pip3 install --user -r temu-audit/requirements.txt
```

## 4. 配置 Web 应用
- 点顶部 Web 标签 → Add a new web app
- 选 Manual configuration → Python 3.10
- 找到 WSGI configuration file 链接，点进去
- **全部删除**，替换为 pa_wsgi.py 的内容
- 把文件中的 `{USERNAME}` 改成你的 PythonAnywhere 用户名
- 把 Stripe 密钥替换为真实值

## 5. 启动
回到 Web 标签页，点绿色 Reload 按钮。

访问 https://你的用户名.pythonanywhere.com 即可。
