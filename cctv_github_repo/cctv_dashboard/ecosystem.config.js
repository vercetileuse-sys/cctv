// PM2 Ecosystem Config — run: pm2 start ecosystem.config.js
module.exports = {
  apps: [{
    name: 'cctv-dashboard',
    script: 'app.py',
    interpreter: 'python',
    cwd: __dirname,
    instances: 1,
    autorestart: true,
    watch: false,
    max_memory_restart: '512M',
    env: {
      FLASK_ENV: 'production',
      PYTHONUNBUFFERED: '1'
    },
    error_file: './logs/error.log',
    out_file: './logs/out.log',
    log_date_format: 'YYYY-MM-DD HH:mm:ss',
    merge_logs: true,
  }]
};
