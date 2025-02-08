FROM mcr.microsoft.com/playwright:focal
WORKDIR /app
COPY . .
RUN rm package-lock.json
RUN npm install
RUN npx playwright install  
RUN npm run build
ENTRYPOINT ["npm", "run"]
CMD ["test"]
# 用Docker运行项目
# 1.构建Docker镜像：docker build -t my-node-app .
# 2.运行Docker容器：docker run -p 8788:8788 -d --name docsify my-node-app
# node app.js
# 3.查看Docker容器：docker ps
# 4.查看Docker日志：docker logs <container id>
# 5.停止Docker容器：docker stop <container id>
# 6.删除Docker容器：docker rm <container id>
# 7.删除Docker镜像：docker rmi <image id>