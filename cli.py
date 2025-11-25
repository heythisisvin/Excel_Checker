"""cli.py
Simple CLI to glue the tools together.
"""
import argparse
import json
from basic_corruption_checker import check_xlsx_structure
from analyzer import analyze_xlsx
from report_generator import generate_report

def main():
    p = argparse.ArgumentParser(description='Excel scanner tools')
    p.add_argument('file', help='xlsx file to scan')
    p.add_argument('--check', action='store_true', help='Run quick corruption check')
    p.add_argument('--analyze', action='store_true', help='Run deep analyzer')
    p.add_argument('--report', help='Write HTML report path (run analyze first)')
    args = p.parse_args()

    if args.check:
        ok, msg = check_xlsx_structure(args.file)
        print('OK' if ok else 'PROBLEM', '-', msg)

    if args.analyze:
        res = analyze_xlsx(args.file)
        print(json.dumps(res, indent=2))
        # save analyzer result to file for report if requested
        if args.report:
            with open(args.report + '.json', 'w', encoding='utf-8') as f:
                json.dump(res, f, indent=2)
            generate_report(res, args.report)
            print('Report generated at', args.report)

if __name__ == '__main__':
    main()
