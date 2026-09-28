"""Atomic invocation summaries; bundle journals remain delivery authority."""
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4
from .report_artifacts import write_json


def timestamp():
    return datetime.now(timezone.utc).isoformat()


class RunRecord:
    def __init__(self, root, report_date, **context):
        self.root = Path(root)
        run_id = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S') + '-' + uuid4().hex[:12]
        self.path = self.root / 'invocations' / (run_id + '.json')
        self.value = {'schema_version': 1, 'run_id': run_id, 'date': str(report_date),
                      'started_at': timestamp(), 'status': 'running', 'stage': 'configuration',
                      'scheduled': False, 'results': [], **context}

    def save(self, *, publish=True):
        self.value['updated_at'] = timestamp()
        write_json(self.path, self.value)
        if publish:
            write_json(self.root / 'last-run.json', self.value)
            if self.value['scheduled']:
                write_json(self.root / 'last-report-run.json', self.value)

    def finish(self, status, *, publish=True, **details):
        self.value.update(status=status, finished_at=timestamp(), **details)
        self.save(publish=publish)
