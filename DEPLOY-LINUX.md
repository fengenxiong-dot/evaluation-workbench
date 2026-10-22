# Linux 部署说明

当前版本：`0.3.0`

## 推荐环境

- 64 位 Linux
- Docker 24 或更高版本
- Docker Compose v2
- 至少 2 GB 可用内存
- 局域网 TCP 3000 端口可访问

## 首次部署

```bash
tar -xzf pingjian-v0.3.0-linux-deploy.tar.gz
cd pingjian-v0.3.0
docker compose up -d --build
```

浏览器访问 `http://服务器局域网IP:3000`。

查看服务状态：

```bash
docker compose ps
docker compose logs --tail=100
```

## 数据位置与备份

所有模板、评价人、任务和评分数据保存在解压目录下的 `server-data` 文件夹。升级或重新构建容器不会删除它。

停止服务后备份：

```bash
docker compose down
tar -czf pingjian-data-backup-$(date +%Y%m%d-%H%M%S).tar.gz server-data
docker compose up -d
```

## 版本升级

每个版本解压到独立目录。升级前先备份旧版本的 `server-data`，再将该文件夹复制到新版本目录：

```bash
cd pingjian-v0.2.0
docker compose down
tar -czf ../pingjian-data-before-v0.3.0.tar.gz server-data
cd ..
cp -a pingjian-v0.2.0/server-data pingjian-v0.3.0/
cd pingjian-v0.3.0
docker compose up -d --build
```

确认新版本正常后再保留或归档旧目录。

## 回滚

如果新版本异常，停止新版本并重新启动旧版本。恢复数据前务必保留当前数据副本，避免覆盖升级后新产生的评价记录。

```bash
cd pingjian-v0.3.0
docker compose down
cd ../pingjian-v0.2.0
docker compose up -d
```

## 不使用 Docker

服务器需安装 Node.js 22 或更高版本：

```bash
npm ci
npm run build
npm run start
```

服务会监听 `0.0.0.0:3000`。
