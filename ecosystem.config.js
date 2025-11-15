// PM2 Ecosystem Configuration for Production

module.exports = {
  apps: [{
    name: 'ca-report-backend',
    script: './server.js',
    instances: 'max',
    exec_mode: 'cluster',
    env: {
      NODE_ENV: 'development',
      PORT: 3000
    },
    env_production: {
      NODE_ENV: 'production',
      PORT: 3000
    },
    error_file: './logs/err.log',
    out_file: './logs/out.log',
    log_file: './logs/combined.log',
    time: true,
    max_memory_restart: '1G',
    autorestart: true,
    watch: false,
    ignore_watch: [
      'node_modules',
      'logs',
      'temp',
      'uploads'
    ]
  }]
};
