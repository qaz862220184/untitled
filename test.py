from flask import Flask
import user
app = Flask(__name__)

@app.route('/')
def hello_world():
    user.main()

if __name__ == '__main__':
    app.run()