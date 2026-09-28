"""Enumerate existing report paths and audiences without doing any report work."""
from pathlib import Path


def selected_reports(cohorts, source_root, audiences=()):
    from ProgramReports.email_program_reports import program_dict
    unknown = set(audiences) - ({'leadership'} | set(program_dict))
    if unknown:
        raise ValueError('Unknown audience; use leadership or a current director dictionary key')
    targets = []
    for cohort in cohorts:
        if not audiences or 'leadership' in audiences:
            targets.append({'cohort': cohort, 'audience': 'leadership',
                            'rolling': Path(source_root) / cohort.overview_template,
                            'programs': None, 'recipients': None})
        for contact, data in program_dict.items():
            if audiences and contact not in audiences:
                continue
            programs = tuple(data['programs'])
            filename = f"{cohort.id}-WeeklyPlacement-{'-'.join(programs)}.xlsx"
            targets.append({'cohort': cohort, 'audience': contact,
                'rolling': Path(source_root) / 'ProgramReports' / f'{cohort.id}_Placement_Reports' / filename,
                'programs': programs, 'recipients': list(data['emails'])})
    return targets
