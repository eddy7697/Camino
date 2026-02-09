FROM nginx:1.27-alpine

# Remove default config
RUN rm /etc/nginx/conf.d/default.conf

# Copy custom nginx config (already listens on 8080)
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Copy static files
COPY index.html   /usr/share/nginx/html/
COPY favicon.svg  /usr/share/nginx/html/
COPY img/         /usr/share/nginx/html/img/

# Non-root setup for K8s security best practices
RUN sed -i '/^pid/d' /etc/nginx/nginx.conf && \
    echo "pid /tmp/nginx.pid;" >> /etc/nginx/nginx.conf && \
    chown -R nginx:nginx /usr/share/nginx/html \
                         /var/cache/nginx \
                         /var/log/nginx \
                         /etc/nginx/conf.d && \
    chmod -R 755 /usr/share/nginx/html

EXPOSE 8080

USER nginx

CMD ["nginx", "-g", "daemon off;"]
