import sys

sys.path.insert(0, r"C:\Users\james\OneDrive - Jones & Wagener\Desktop\Apps\Filedwork app\New folder (33)\venv\Lib\site-packages")

import app

app.init_db()
app.app.run(host="127.0.0.1", port=5017, debug=False, use_reloader=False)
