"""Weekly visualization email text (test pipeline only)."""


def director_greeting(names):
    """Alphabetical: 'Tanya' / 'Amy & Tanya' / 'Amy, Bob & Tanya'."""
    names = sorted(names, key=str.casefold)
    return names[0] if len(names) == 1 else ', '.join(names[:-1]) + ' & ' + names[-1]


def subject_line(class_years, label, report_date):
    """class_years: the report class, e.g. [2027]; several are joined with '&'."""
    years = ' & '.join(str(y) for y in sorted(class_years))
    return f'Class of {years}: {label} Weekly Placement Report {report_date:%Y-%m-%d}'


def body_text(names, report_date):
    return (f'{director_greeting(names)},\n\n'
            f'Here is the weekly placement visualization for {report_date:%m-%d}. '
            'Please let us know if you have any questions.\n\n'
            'Best,\n'
            'BCC Data Team')
