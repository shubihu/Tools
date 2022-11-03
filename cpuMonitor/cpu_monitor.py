from flask import Flask, render_template
from flask_socketio import SocketIO
from threading import Lock
import random
import os
import psutil

async_mode = None
app = Flask(__name__)
app.config['SECRET_KEY'] = 'secret!'
socketio = SocketIO(app, async_mode=async_mode)
thread = None
thread_lock = Lock()

@app.route('/')
def index():
    return render_template('index.html', async_mode=socketio.async_mode)

@socketio.on('connect', namespace='/test_conn')
def test_connect():
    global thread
    with thread_lock:
        if thread is None:
            thread = socketio.start_background_task(target=background_thread)

def background_thread():
    while True:
        socketio.sleep(1)
        # t = random.randint(50, 100)
        l1, l2, l3 = psutil.getloadavg()
        CPU_use = (l3/os.cpu_count()) * 100
        socketio.emit('server_response', {'data': CPU_use}, namespace='/test_conn')

@socketio.on('disconnect', namespace='/chat')
def test_disconnect():
    print('Client disconnected')

if __name__ == '__main__':
    socketio.run(app, host='0.0.0.0', port=5001, debug=True)