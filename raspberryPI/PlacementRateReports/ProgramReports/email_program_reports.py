"""Career-director routing; preparation/delivery is owned by core.pipeline."""
from core.report_artifacts import ROOT, RunOptions

program_dict = {
    "Tracie": {
        "programs": ("BSAcc", "MAcc"),
        "emails": ("Tracie.Laham@byu.edu", "Jonathan.Creer@byu.edu", "bccdata@byu.edu", "quinn_frazier@byu.edu"),
    },
    "Noelani": {
        "programs": ("BSEDM",),
        "emails": ("noelani_wayas@byu.edu", "Jonathan.Creer@byu.edu", "bccdata@byu.edu", "quinn_frazier@byu.edu"),
    },
    "Soraya": {
        "programs": ("BSHRM", "BSStrat"),
        "emails": ("Jonathan.Creer@byu.edu", "Soraya.Cullimore@byu.edu", "quinn_frazier@byu.edu", "bccdata@byu.edu"),
    },
    "Tanya": {
        "programs": ("BSFin",),
        "emails": ("tharmon@byu.edu", "amy@byu.edu", "Jonathan.Creer@byu.edu", "bccdata@byu.edu", "quinn_frazier@byu.edu"),
    },
    "Kurt": {
        "programs": ("BSGSCM",),
        "emails": ("Kurt.Francis@byu.edu", "Jonathan.Creer@byu.edu", "bccdata@byu.edu", "quinn_frazier@byu.edu"),
    },
    "Steven": {
        "programs": ("BSEnt", "BSBusM"),
        "emails": ("Steven.Steele@byu.edu", "Jonathan.Creer@byu.edu", "bccdata@byu.edu", "quinn_frazier@byu.edu"),
    },
    "Bob": {
        "programs": ("BSIS", "MISM"),
        "emails": ("bob.ure@byu.edu","Audrey.Seager@byu.edu", "Jonathan.Creer@byu.edu", "bccdata@byu.edu", "quinn_frazier@byu.edu"),
    },
    "Mike": {
        "programs": ("BSMktg",),
        "emails": ("mike.neuffer@byu.edu", "Jonathan.Creer@byu.edu", "bccdata@byu.edu", "quinn_frazier@byu.edu"),
    },
    "Perry": {
        "programs": ("MBA",),
        "emails": ("perry.christensen@byu.edu", "Jonathan.Creer@byu.edu", "bccdata@byu.edu", "quinn_frazier@byu.edu"),
    },
    "Staci": {
        "programs": ("MPA",),
        "emails": ("staci_carroll@byu.edu", "Jonathan.Creer@byu.edu", "bccdata@byu.edu", "quinn_frazier@byu.edu"),
    },
    "Erica": {
        "programs": ("Healthcare",),
        "emails": ("erica_card@byu.edu", "Jonathan.Creer@byu.edu", "bccdata@byu.edu", "quinn_frazier@byu.edu"),
    },
}


def mainflow(cohort, a, today, options=None):
    from core.pipeline import execute_bundle
    options = options or RunOptions()
    results = []
    for contact, data in program_dict.items():
        if options.audiences and contact not in options.audiences:
            continue
        programs = tuple(data['programs'])
        filename = f"{cohort.id}-WeeklyPlacement-{'-'.join(programs)}.xlsx"
        source = options.source_root / 'ProgramReports' / f'{cohort.id}_Placement_Reports' / filename
        results.append(execute_bundle(cohort, a, today, contact, source, programs,
                                      list(data['emails']), options))
    return results
