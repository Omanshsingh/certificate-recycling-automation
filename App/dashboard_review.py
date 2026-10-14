"""Coordinate reminder review with one live dashboard tab; no file changes."""
import threading
import time
import uuid


class DashboardReview:
    def __init__(self):
        self.clients = {}
        self.pending = None
        self.lock = threading.Lock()

    def presence(self, client_id, visible=False, acknowledged=None):
        if not isinstance(client_id, str) or not 1 <= len(client_id) <= 64:
            raise ValueError('Invalid dashboard tab identifier.')
        with self.lock:
            now = time.monotonic()
            self.clients = {key: value for key, value in self.clients.items() if now - value['seen'] < 15}
            self.clients[client_id] = {'seen': now, 'visible': bool(visible)}
            request = self.pending
            if request and request['client_id'] == client_id:
                if acknowledged == request['id']:
                    request['acknowledged'] = True
                elif not request['acknowledged']:
                    return {'id': request['id']}
            return None

    def request(self):
        with self.lock:
            now = time.monotonic()
            live = [(key, value) for key, value in self.clients.items() if now - value['seen'] < 15]
            if not live:
                return {'reused': False}
            client_id, _ = max(live, key=lambda item: (item[1]['visible'], item[1]['seen']))
            self.pending = {'id': uuid.uuid4().hex, 'client_id': client_id, 'acknowledged': False}
            return {'reused': True, 'request_id': self.pending['id']}

    def status(self, request_id, cancel=False):
        with self.lock:
            request = self.pending
            accepted = bool(request and request['id'] == request_id and request['acknowledged'])
            if cancel and request and request['id'] == request_id and not accepted:
                self.pending = None
            return {'acknowledged': accepted}
