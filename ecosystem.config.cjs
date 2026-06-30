const path = require("path");

const rootDir = __dirname;

module.exports = {
  apps: [
    {
      name: "traffic-api",
      cwd: path.join(rootDir, "01_Program", "09_else"),
      script: path.join(rootDir, ".venv", "Scripts", "python.exe"),
      args: "-m traffic_api.main",
      interpreter: "none",
      out_file: path.join(rootDir, "02_Result", "traffic-api-pm2-out.log"),
      error_file: path.join(rootDir, "02_Result", "traffic-api-pm2-error.log"),
      merge_logs: true,
      autorestart: true,
      env: {
        PYTHONUNBUFFERED: "1",
      },
    },
  ],
};
