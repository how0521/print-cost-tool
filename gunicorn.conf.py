import os
bind = "0.0.0.0:{}".format(os.environ.get("PORT", "8080"))
timeout = 300
workers = 1
