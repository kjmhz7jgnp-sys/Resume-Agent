# Resume Agent 试用版部署包

这是一个可部署到 Render、Railway 或 Fly.io 的简历适配工具。

## 功能

- 上传 `.docx`、`.pdf`、`.txt` 原简历
- 根据 JD、候选人要求、公司名称、岗位和公司类型生成适配版 Word
- 在简历最上方插入 3-5 条综合评价
- 设置 `APP_PASSWORD` 后启用浏览器基础密码保护
- 上传文件和生成文件只在临时目录处理，请求结束后删除

## 环境变量

必填：

- `APP_PASSWORD`: 访问密码

可选：

- `APP_USERNAME`: 访问账号，默认 `admin`
- `PORT`: 云平台通常自动提供

## Render

1. 新建 Web Service，连接包含本目录的 Git 仓库。
2. Root Directory 设为 `deploy/resume-agent`。
3. Build Command: `pip install -r requirements.txt`
4. Start Command: `python server.py`
5. 添加环境变量 `APP_PASSWORD`。

如果使用 Blueprint，可直接用 `render.yaml`。

## Railway

1. 新建项目并连接 Git 仓库。
2. Root Directory 设为 `deploy/resume-agent`。
3. 添加环境变量 `APP_PASSWORD`。
4. Railway 会读取 `railway.json`，启动命令为 `python server.py`。

## Fly.io

1. 在本目录执行 `fly launch`。
2. 设置密码：

```bash
fly secrets set APP_PASSWORD="your-password"
```

3. 部署：

```bash
fly deploy
```

## 访问

部署完成后打开平台给出的 HTTPS 地址。浏览器会弹出账号密码框：

- 用户名：`admin`，或你设置的 `APP_USERNAME`
- 密码：你设置的 `APP_PASSWORD`
