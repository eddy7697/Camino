# 朝聖之路 Camino de Santiago

32 天、800 公里——從法國巴黎到西班牙聖地牙哥，再到世界盡頭菲斯特雷角的信仰旅程紀錄。

## 專案結構

```
├── index.html        # 靜態網頁（主要展示頁面）
├── img/              # 朝聖之路沿途照片（01.jpg ~ 30.jpg）
├── create_pptx.py    # PowerPoint 簡報產生器
├── note.md           # 圖片註解
├── Dockerfile        # Docker 容器化設定
├── nginx.conf        # Nginx 設定檔（port 8080）
└── .dockerignore     # Docker build 排除清單
```

## 快速開始

### 直接開啟

用瀏覽器直接開啟 `index.html` 即可瀏覽。

### Docker 部署

#### 建置映像檔

```bash
docker build -t camino-website .
```

#### 啟動容器

```bash
docker run -d --name camino-website -p 8080:8080 camino-website
```

啟動後開啟瀏覽器訪問：http://localhost:8080

#### 停止與移除容器

```bash
# 停止
docker stop camino-website

# 移除
docker rm camino-website
```

#### 健康檢查

容器內建 `/healthz` 端點，供 K8s liveness/readiness probe 使用：

```bash
curl http://localhost:8080/healthz
```

### Kubernetes 部署

```yaml
apiVersion: apps/v1
kind: Deployment
metadata:
  name: camino-website
spec:
  replicas: 2
  selector:
    matchLabels:
      app: camino-website
  template:
    metadata:
      labels:
        app: camino-website
    spec:
      containers:
      - name: camino-website
        image: <your-registry>/camino-website:latest
        ports:
        - containerPort: 8080
        livenessProbe:
          httpGet:
            path: /healthz
            port: 8080
        readinessProbe:
          httpGet:
            path: /healthz
            port: 8080
---
apiVersion: v1
kind: Service
metadata:
  name: camino-website
spec:
  selector:
    app: camino-website
  ports:
  - port: 80
    targetPort: 8080
```

## 產生 PowerPoint 簡報

需要 Python 3 環境：

```bash
pip install python-pptx Pillow lxml
python create_pptx.py
```

執行後會在專案目錄產生 `朝聖之路.pptx`。
