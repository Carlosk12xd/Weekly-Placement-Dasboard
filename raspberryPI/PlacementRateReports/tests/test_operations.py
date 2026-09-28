"""Operational regressions. SQL, SMTP and dotenv parsing are mocked explicitly."""
from contextlib import redirect_stdout
from datetime import date
import fcntl
from io import StringIO
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import Mock, patch
import yaml

from core import daily_runner, database, delivery, doctor, pipeline, status
from core.report_artifacts import ROOT, RunOptions, sha256, write_json
from Visualization.build_visualization import load_profile

FIXTURES = ROOT / 'tests/fixtures'
AS_OF = date(2026, 8, 21)


class AcceptingSMTP:
    sent = []
    passwords = []
    def __init__(self, *args, **kwargs): pass
    def __enter__(self): return self
    def __exit__(self, *args): pass
    def ehlo(self): pass
    def starttls(self, **kwargs): pass
    def login(self, sender, password): self.passwords.append(password)
    def send_message(self, message, **kwargs):
        self.sent.append((message, kwargs))
        return {}


class OperationsTests(TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.work = Path(self.temp.name)
        self.run_root = self.work / 'runs'
        self.config_path = ROOT / 'config/cohorts.yml'
        self.cohort = next(c for c in daily_runner.load_cohorts()[1] if c.id == '2027')
        self.finance = FIXTURES / 'ProgramReports/2027_Placement_Reports/2027-WeeklyPlacement-BSFin.xlsx'
        AcceptingSMTP.sent = []
        AcceptingSMTP.passwords = []

    def tearDown(self):
        self.temp.cleanup()

    def run_offline(self, *extra, source=FIXTURES, run_date=AS_OF):
        with patch.object(daily_runner, 'ROOT', self.work / 'project'), \
             patch('smtplib.SMTP', side_effect=AssertionError('Offline must not call SMTP')), \
             patch.object(database, 'connect_database', side_effect=AssertionError('Offline must not call SQL')):
            return daily_runner.main(['--config', str(self.config_path), '--offline', '--cohort', '2027',
                '--date', str(run_date), '--source-root', str(source), '--run-root', str(self.run_root), *extra])

    def test_directory_settings_are_isolated_and_process_values_take_precedence(self):
        leader, program = self.work / 'LeadershipReport', self.work / 'ProgramReports'
        leader.mkdir()
        program.mkdir()
        (leader / '.env').write_text('SMTP_PASS=leader-mock\nDB_NAME=leader-db\nTO_ADDRS=leader@example.invalid\n')
        (program / '.env').write_text('SMTP_PASS=program-mock\nDB_NAME=program-db\n')
        def parse(path):
            return dict(line.split('=', 1) for line in path.read_text().splitlines())
        # We test our settings routing; python-dotenv syntax itself is not under test.
        parser = Mock(side_effect=parse)
        with patch.dict(sys.modules, {'dotenv': SimpleNamespace(dotenv_values=parser)}), \
             patch.dict(os.environ, {'DB_HOST': 'inherited-host'}, clear=True):
            before = dict(os.environ)
            a = database.load_environment(leader)
            b = database.load_environment(program)
            self.assertEqual(a['SMTP_PASS'], 'leader-mock')
            self.assertEqual(b['SMTP_PASS'], 'program-mock')
            self.assertNotIn('TO_ADDRS', b)
            self.assertEqual(b['DB_HOST'], 'inherited-host')
            self.assertEqual(dict(os.environ), before)
            os.environ['DB_NAME'] = 'explicit-process-db'
            self.assertEqual(database.load_environment(program)['DB_NAME'], 'explicit-process-db')
        self.assertEqual(parser.call_args_list[0].args, (leader / '.env',))
        self.assertEqual(parser.call_args_list[1].args, (program / '.env',))

    def test_database_uses_returned_settings_not_shared_environment(self):
        connect = Mock(return_value='mock-connection')
        connector = SimpleNamespace(connect=connect)
        mysql = SimpleNamespace(connector=connector)
        settings = {'DB_HOST': 'local', 'DB_USER': 'user', 'DB_PASSWORD': 'private-mock', 'DB_NAME': 'program'}
        with patch.object(database, 'load_environment', return_value=settings), \
             patch.dict(sys.modules, {'mysql': mysql, 'mysql.connector': connector}), \
             patch.dict(os.environ, {'DB_NAME': 'wrong'}):
            self.assertEqual(database.connect_database(self.work), 'mock-connection')
        self.assertEqual(connect.call_args.kwargs['database'], 'program')
        self.assertEqual(connect.call_args.kwargs['password'], 'private-mock')

    def test_unsafe_yaml_types_and_duplicate_cohorts_are_rejected(self):
        baseline = yaml.safe_load(self.config_path.read_text())
        mutations = [lambda c: c.update(delivery_enabled='false'),
            lambda c: c.update(tracking_grace_period_days=-1),
            lambda c: c['cohorts'][0].update(enabled='false'),
            lambda c: c['cohorts'][0].update(human_modes=[True]),
            lambda c: c['cohorts'].append(dict(c['cohorts'][0])),
            lambda c: c['cohorts'][0].update(start_date='2028-01-01')]
        path = self.work / 'cohorts.yml'
        for mutate in mutations:
            with self.subTest(mutation=mutate):
                config = json.loads(json.dumps(baseline))
                mutate(config)
                path.write_text(yaml.safe_dump(config))
                with self.assertRaises(ValueError):
                    daily_runner.load_cohorts(path)
        (self.work / 'config').mkdir()
        (self.work / 'config/visualizations.yml').write_text('enabled: "false"\ncohorts: [2027]\n')
        with self.assertRaises(ValueError):
            load_profile(self.work)

    def test_early_configuration_failure_replaces_stale_success_summary(self):
        write_json(self.run_root / 'last-run.json', {'status': 'completed', 'date': 'old', 'results': []})
        bad = self.work / 'bad.yml'
        bad.write_text('delivery_enabled: "false"\ncohorts: []\n')
        self.config_path = bad
        self.assertEqual(self.run_offline('--audience', 'Erica'), 1)
        saved = json.loads((self.run_root / 'last-run.json').read_text())
        self.assertEqual(saved['status'], 'failed')
        self.assertEqual(saved['date'], str(AS_OF))
        self.assertIn('delivery_enabled', saved['error'])
        self.assertIn('finished_at', saved)
        self.assertEqual(len(list((self.run_root / 'invocations').glob('*.json'))), 1)

    def test_missing_workbook_does_not_stop_other_audiences(self):
        source = self.work / 'source'
        fixture = FIXTURES / 'ProgramReports/2027_Placement_Reports/2027-WeeklyPlacement-Healthcare.xlsx'
        target = source / fixture.relative_to(FIXTURES)
        target.parent.mkdir(parents=True)
        shutil.copy2(fixture, target)
        self.assertEqual(self.run_offline('--audience', 'Tanya', '--audience', 'Erica', source=source), 1)
        record = json.loads((self.run_root / 'last-run.json').read_text())
        self.assertEqual([(r['audience'], r['status']) for r in record['results']], [('Tanya', 'failed'), ('Erica', 'preview')])
        self.assertEqual(record['expected_bundles'], 2)
        self.assertTrue((self.run_root / 'last-report-run.json').is_file())

    def test_corrupt_bundle_is_preserved_and_next_audience_still_runs(self):
        path = self.run_root / f'preview/{AS_OF}/2027/Tanya/manifest.json'
        path.parent.mkdir(parents=True)
        path.write_text('{incomplete-ledger')
        self.assertEqual(self.run_offline('--audience', 'Tanya', '--audience', 'Erica'), 1)
        self.assertEqual(path.read_text(), '{incomplete-ledger')
        record = json.loads((self.run_root / 'last-run.json').read_text())
        self.assertEqual(record['results'][1]['status'], 'preview')

    def test_missing_converter_holds_pdf_mail_without_blocking_healthcare_or_archive_previews(self):
        with patch('Visualization.render_excel.dependencies', side_effect=RuntimeError('Missing renderer executables')):
            code = self.run_offline('--audience', 'Tanya', '--audience', 'Erica')
        self.assertEqual(code, 1)
        record = json.loads((self.run_root / 'last-run.json').read_text())
        self.assertEqual(record['results'][0]['routes'], {'box': 'prepared'})
        self.assertEqual(record['results'][0]['human_held'], 'Required PDF unavailable')
        self.assertEqual(record['results'][1]['status'], 'preview')
        self.assertEqual(record['results'][1]['routes']['human'], 'prepared')

    def test_live_converter_failure_sends_only_archive_using_group_settings(self):
        rolling = self.work / self.finance.name
        shutil.copy2(self.finance, rolling)
        options = RunOptions(dry_run=False, run_root=self.run_root, source_root=FIXTURES)
        fake_date = Mock()
        fake_date.today.return_value = AS_OF
        settings = {'SMTP_PASS': 'group-mock'}
        with patch.object(pipeline, 'date', fake_date), \
             patch('ProgramReports.create_program_reports.main'), patch('ProgramReports.placement_details.main'), \
             patch.object(database, 'load_environment', return_value=settings), \
             patch.object(pipeline, 'render_pdf', side_effect=RuntimeError('missing converter')), \
             patch('smtplib.SMTP', AcceptingSMTP), patch.dict(os.environ, {'SMTP_PASS': 'wrong-global'}):
            result = pipeline.execute_bundle(self.cohort, 0, AS_OF, 'Tanya', rolling,
                                            ('BSFin',), ['test@example.invalid'], options)
        self.assertEqual(result['status'], 'failed')
        self.assertEqual(result['routes'], {'box': 'sent'})
        self.assertEqual(len(AcceptingSMTP.sent), 1)
        self.assertEqual(AcceptingSMTP.passwords, ['group-mock'])
        self.assertEqual(len(list(AcceptingSMTP.sent[0][0].iter_attachments())), 1)

    def test_lock_contention_records_blocked_attempt_without_overwriting_active_summary(self):
        lock_path = self.work / 'project/runs/runner.lock'
        lock_path.parent.mkdir(parents=True)
        write_json(self.run_root / 'last-run.json', {'status': 'running', 'run_id': 'active', 'results': []})
        original = (self.run_root / 'last-run.json').read_bytes()
        with lock_path.open('a') as lock:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
            self.assertEqual(self.run_offline('--audience', 'Erica'), 1)
        self.assertEqual((self.run_root / 'last-run.json').read_bytes(), original)
        blocked = json.loads(next((self.run_root / 'invocations').glob('*.json')).read_text())
        self.assertEqual(blocked['status'], 'blocked')

    def test_skipped_day_preserves_last_report_and_interrupt_is_recorded(self):
        write_json(self.run_root / 'last-report-run.json', {'status': 'completed', 'results': []})
        original = (self.run_root / 'last-report-run.json').read_bytes()
        self.assertEqual(self.run_offline(run_date=date(2026, 8, 20)), 0)
        self.assertEqual((self.run_root / 'last-report-run.json').read_bytes(), original)
        self.assertEqual(json.loads((self.run_root / 'last-run.json').read_text())['status'], 'skipped')
        with patch.object(pipeline, 'execute_bundle', side_effect=KeyboardInterrupt):
            self.assertEqual(self.run_offline('--audience', 'Erica'), 130)
        record = json.loads((self.run_root / 'last-run.json').read_text())
        self.assertEqual(record['status'], 'interrupted')
        self.assertEqual(record['current_bundle'], '2027/Erica')

    def test_readiness_is_read_only_and_offline_skips_credentials(self):
        with patch.object(database, 'load_environment', side_effect=AssertionError('Offline read settings')), \
             patch('socket.create_connection', side_effect=AssertionError('Readiness opened network')), \
             patch.object(doctor.shutil, 'which', return_value='/mock/bin/tool'):
            result = doctor.inspect_setup(source_root=FIXTURES, run_root=self.run_root,
                report_date=AS_OF, cohort_ids=['2027'], offline=True)
        self.assertEqual(result['status'], 'passed', result)
        self.assertFalse(self.run_root.exists())
        self.assertEqual(len([c for c in result['checks'] if c['check'].startswith('input:')]), 12)
        self.assertFalse(result['network_checked'])

    def test_readiness_reports_missing_renderer_only_for_supported_audiences(self):
        with patch.object(doctor.shutil, 'which', return_value=None):
            finance = doctor.inspect_setup(source_root=FIXTURES, run_root=self.run_root,
                report_date=AS_OF, cohort_ids=['2027'], audiences=['Tanya'], offline=True)
            healthcare = doctor.inspect_setup(source_root=FIXTURES, run_root=self.run_root,
                report_date=AS_OF, cohort_ids=['2027'], audiences=['Erica'], offline=True)
        self.assertEqual(finance['status'], 'failed')
        self.assertEqual(healthcare['status'], 'passed')

    def test_readiness_does_not_display_setting_values(self):
        secret = 'never-display-this-mock-password'
        settings = {'DB_HOST': 'mock-host', 'DB_USER': 'mock-user', 'DB_PASSWORD': secret,
                    'DB_NAME': 'mock-name', 'SMTP_PASS': secret, 'TO_ADDRS': 'leader@example.invalid',
                    'Zoho_Box': 'archive@example.invalid'}
        with patch.object(database, 'load_environment', return_value=settings):
            result = doctor.inspect_setup(source_root=FIXTURES, run_root=self.run_root,
                report_date=AS_OF, cohort_ids=['2027'], audiences=['leadership'])
        encoded = json.dumps(result)
        self.assertNotIn(secret, encoded)
        self.assertNotIn('leader@example.invalid', encoded)
        self.assertNotIn('mock-host', encoded)

    def test_status_detects_corrupt_artifacts_and_uncertain_delivery_without_sending(self):
        self.assertEqual(self.run_offline('--audience', 'Erica'), 0)
        with patch('smtplib.SMTP', side_effect=AssertionError('Status must not send')):
            report = status.inspect_run(self.run_root, verify=True)
            self.assertEqual(report['status'], 'completed')
            self.assertEqual(report['bundles'][0]['hashes'], 'verified')
            directory = self.run_root / f'preview/{AS_OF}/2027/Erica'
            manifest_path = directory / 'manifest.json'
            manifest = json.loads(manifest_path.read_text())
            manifest['routes']['human']['state'] = 'uncertain'
            write_json(manifest_path, manifest)
            report = status.inspect_run(self.run_root)
            self.assertEqual(report['status'], 'failed')
            self.assertIn('SMTP outcome', report['bundles'][0]['error'])
            manifest['routes']['human']['state'] = 'prepared'
            write_json(manifest_path, manifest)
            (directory / manifest['artifacts']['xlsx']['file']).write_bytes(b'corrupt')
            report = status.inspect_run(self.run_root, verify=True)
            self.assertEqual(report['status'], 'failed')
            self.assertIn('Artifact missing or changed', report['bundles'][0]['error'])

    def test_invalid_envelope_fails_before_smtp(self):
        attachment = self.work / 'sample.xlsx'
        attachment.write_bytes(b'fixture')
        manifest = {'routes': {}}
        msg = delivery.message('Preview', 'Body', [attachment], ['valid@example.invalid'], message_key='address-test')
        delivery.save_route(self.work, manifest, 'human', msg, ['invalid@@example.invalid'])
        with patch('smtplib.SMTP', side_effect=AssertionError('Invalid envelope reached SMTP')):
            with self.assertRaises(ValueError):
                delivery.deliver(self.work, manifest, dry_run=False, settings={'SMTP_PASS': 'mock'})
        self.assertEqual(manifest['routes']['human']['state'], 'prepared')
