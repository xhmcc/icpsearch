from PyQt5.QtCore import QObject, pyqtSignal

class ContinueQuerySignal(QObject):
    ask_continue = pyqtSignal(str)
    continue_result = pyqtSignal(str)

continue_signal = ContinueQuerySignal()

_continue_result = None

def should_continue_callback(msg):
    from PyQt5.QtCore import QEventLoop
    global _continue_result
    _continue_result = None
    def on_result(res):
        global _continue_result
        _continue_result = res
        loop.quit()
    continue_signal.continue_result.connect(on_result)
    continue_signal.ask_continue.emit(msg)
    loop = QEventLoop()
    while _continue_result is None:
        loop.processEvents()
    continue_signal.continue_result.disconnect(on_result)
    return _continue_result 