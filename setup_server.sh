#!/bin/bash
# IBK Dashboard - VPS sozlash skripti (Ubuntu 22.04)
# Foydalanish: bash setup_server.sh

set -e

echo "=== IBK Dashboard Server sozlanmoqda ==="

# Python va kerakli paketlar
apt-get update -y
apt-get install -y python3 python3-pip python3-venv git nginx

# Virtual environment
python3 -m venv /opt/ibk-venv
source /opt/ibk-venv/bin/activate
pip install -r /opt/ibk-dashboard/requirements.txt

# systemd service yaratish
cat > /etc/systemd/system/ibk-dashboard.service << 'EOF'
[Unit]
Description=IBK Dashboard Server
After=network.target

[Service]
Type=simple
User=root
WorkingDirectory=/opt/ibk-dashboard
ExecStart=/opt/ibk-venv/bin/python3 ibk_dashboard.py
Restart=always
RestartSec=5
Environment=IBK_HOST=0.0.0.0
Environment=IBK_PORT=8788
Environment=IBK_GENERATOR_DIR=/opt/ibk-dashboard

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable ibk-dashboard
systemctl start ibk-dashboard

# Nginx sozlash (port 80 -> 8788)
cat > /etc/nginx/sites-available/ibk-dashboard << 'EOF'
server {
    listen 80;
    server_name _;

    client_max_body_size 200M;

    location / {
        proxy_pass http://127.0.0.1:8788;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_read_timeout 300s;
    }
}
EOF

ln -sf /etc/nginx/sites-available/ibk-dashboard /etc/nginx/sites-enabled/
rm -f /etc/nginx/sites-enabled/default
nginx -t && systemctl restart nginx

# Firewall
ufw allow 22
ufw allow 80
ufw allow 443
ufw --force enable

echo ""
echo "=== Tayyor! ==="
echo "Saytga kirish: http://$(curl -s ifconfig.me)"
