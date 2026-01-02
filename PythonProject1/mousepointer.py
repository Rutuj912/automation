from pynput import *

def get_coords(x,y):
    print("now at: {}". format((x,y)))

with mouse.Listener(on_move=get_coords) as listener:
    listener.join()

