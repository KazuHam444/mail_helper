"""Repeatable synthetic email evaluation. --ai uses the configured API; never sends real mail."""
import argparse
from concurrent.futures import ThreadPoolExecutor
import json
from pathlib import Path
import hr_digest as app


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--ai', action='store_true')
    parser.add_argument('--split', choices=['all', 'development', 'holdout'], default='all')
    parser.add_argument('--output')
    args = parser.parse_args()
    if args.ai:
        from dotenv import load_dotenv
        load_dotenv(app.ROOT / '.env')
        if not app.ai_enabled():
            raise SystemExit('AI is not configured.')
    cases = json.loads((app.ROOT / 'tests/hr_eval_cases.json').read_text(encoding='utf-8'))
    cases = [c for c in cases if args.split == 'all' or c['split'] == args.split]
    def check(case):
        decision, final = app.classify(case) if args.ai else (app.rules(case)[0], True)
        return dict(id=case['id'], expected=case['expected'], actual=decision['notify'],
                    correct=case['expected'] == decision['notify'], final=final, source=decision['source'])
    with ThreadPoolExecutor(max_workers=3 if args.ai else 1) as pool:
        results = list(pool.map(check, cases))
    summary = dict(total=len(results), missed=sum(r['expected'] and not r['actual'] for r in results),
                   false_alerts=sum(not r['expected'] and r['actual'] for r in results),
                   ai_calls=sum(r['source'] in ('ai', 'rules-confirmed-receipt', 'rules-confirmed-action') for r in results),
                   fallback=sum(not r['final'] for r in results))
    print(json.dumps(summary))
    print('Failed cases: ' + ', '.join(r['id'] for r in results if not r['correct']))
    if args.output:
        Path(args.output).write_text(json.dumps(dict(summary=summary, results=results), indent=2), encoding='utf-8')
    return 0 if all(r['correct'] and r['final'] for r in results) else 1


if __name__ == '__main__':
    raise SystemExit(main())
