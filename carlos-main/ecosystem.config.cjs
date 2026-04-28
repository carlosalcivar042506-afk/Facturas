module.exports = {
  apps: [
    {
      name: 'factura-ocr',
      script: '/usr/local/bin/python3',
      args: '/home/user/factura-ocr/server.py',
      cwd: '/home/user/factura-ocr',
      watch: false,
      instances: 1,
      exec_mode: 'fork',
      env: {
        PYTHONUNBUFFERED: '1'
      }
    }
  ]
}
