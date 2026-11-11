"""Exercise Tk startup and default layout without user interaction."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from filter_gui import App

app = App()
app.update_idletasks()
try:
    assert app.winfo_width() >= app.winfo_reqwidth(), 'Default window clips content horizontally'
    assert app.winfo_height() >= app.winfo_reqheight(), 'Default window clips content vertically'
    assert str(app.run_button['state']) == 'normal'
    assert app.src.get() == ''
finally:
    app.destroy()
print('Desktop UI startup and default layout verified.')
