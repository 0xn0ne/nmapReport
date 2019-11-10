#!/usr/bin/env python

from libnmap.parser import NmapParser, NmapParserException
from xlsxwriter import Workbook
from datetime import datetime
import os.path


def generate_summary(workbook, sheet, report):
    summary_header = ['Source', 'Command', 'Version', 'Scan Type', 'Started', 'Completed', 'Hosts Total', 'Hosts Up',
                      'Hosts Down']
    summary_body = {
        'Source': lambda report: parsed.source,
        'Command': lambda report: report.commandline,
        'Version': lambda report: report.version,
        'Scan Type': lambda report: report.scan_type,
        'Started': lambda report: datetime.utcfromtimestamp(report.started).strftime('%Y-%m-%d %H:%M:%S (UTC)'),
        'Completed': lambda report: datetime.utcfromtimestamp(report.endtime).strftime('%Y-%m-%d %H:%M:%S (UTC)'),
        'Hosts Total': lambda report: report.hosts_total,
        'Hosts Up': lambda report: report.hosts_up,
        'Hosts Down': lambda report: report.hosts_down
    }

    for idx, item in enumerate(summary_header):
        sheet.write(0, idx, item, workbook.custom_formats['fmt_bold'])
        for idx, item in enumerate(summary_header):
            sheet.write(sheet.lastrow + 1, idx, summary_body[item](report))

    sheet.lastrow = sheet.lastrow + 1


def generate_hosts(workbook, sheet, report):
    sheet.autofilter('A1:E1')
    sheet.freeze_panes(1, 0)

    hosts_header = ['Host', 'IP', 'Status', 'Services', 'OS']
    hosts_body = {
        'Host': lambda host: next(iter(host.hostnames), ''),
        'IP': lambda host: host.address,
        'Status': lambda host: host.status,
        'Services': lambda host: len(host.services),
        'OS': lambda host: os_class_string(host.os_class_probabilities())
    }

    for idx, item in enumerate(hosts_header):
        sheet.write(0, idx, item, workbook.custom_formats['fmt_bold'])

    row = sheet.lastrow
    for host in report.hosts:
        for idx, item in enumerate(hosts_header):
            sheet.write(row + 1, idx, hosts_body[item](host))
        row += 1

    sheet.lastrow = row


def generate_results(workbook, sheet, report):
    sheet.autofilter('A1:N1')
    sheet.freeze_panes(1, 0)

    sheet.data_validation('N2:N$1048576', {
        'validate': 'list',
        'source': ['Y', 'N', 'N/A']}
                          )

    results_header = ['Host', 'IP', 'Port', 'Protocol', 'Status', 'Service', 'Tunnel', 'Method', 'Confidence', 'Reason',
                      'Product', 'Version', 'Extra', 'Flagged', 'Notes']
    results_body = {
        'Host': lambda host, service: next(iter(host.hostnames), ''),
        'IP': lambda host, service: host.address,
        'Port': lambda host, service: service.port,
        'Protocol': lambda host, service: service.protocol,
        'Status': lambda host, service: service.state,
        'Service': lambda host, service: service.service,
        'Tunnel': lambda host, service: service.tunnel,
        'Method': lambda host, service: service.service_dict.get('method', ''),
        'Confidence': lambda host, service: float(service.service_dict.get('conf', '0')) / 10,
        'Reason': lambda host, service: service.reason,
        'Product': lambda host, service: service.service_dict.get('product', ''),
        'Version': lambda host, service: service.service_dict.get('version', ''),
        'Extra': lambda host, service: service.service_dict.get('extrainfo', ''),
        'Flagged': lambda host, service: 'N/A',
        'Notes': lambda host, service: ''
    }

    results_format = {'Confidence': workbook.custom_formats['fmt_conf']}

    print('[+] Processing {}'.format(report.summary))
    for idx, item in enumerate(results_header):
        sheet.write(0, idx, item, workbook.custom_formats['fmt_bold'])

    row = sheet.lastrow
    for host in report.hosts:
        print('[+] Processing {}'.format(host))
        for service in host.services:
            for idx, item in enumerate(results_header):
                sheet.write(row + 1, idx, results_body[item](host, service), results_format.get(item, None))
            row += 1

    sheet.lastrow = row


def setup_workbook_formats(workbook):
    formats = {
        'fmt_bold': workbook.add_format({'bold': True}),
        'fmt_conf': workbook.add_format()
    }

    formats['fmt_conf'].set_num_format('0%')
    return formats


def os_class_string(os_class_array):
    return ' | '.join(['{0} ({1}%)'.format(os_string(osc), osc.accuracy) for osc in os_class_array])


def os_string(os_class):
    rval = '{0}, {1}'.format(os_class.vendor, os_class.osfamily)
    if len(os_class.osgen):
        rval += '({0})'.format(os_class.osgen)
    return rval


def main(reports, workbook):
    sheets = {
        'Summary': generate_summary,
        'Hosts': generate_hosts,
        'Results': generate_results
    }

    setattr(workbook, 'custom_formats', setup_workbook_formats(workbook))

    for sheet_name, sheet_func in sheets.items():
        sheet = workbook.add_worksheet(sheet_name)
        sheet.lastrow = 0
        for report in reports:
            sheet_func(workbook, sheet, report)
    workbook.close()


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument('-r', '--reports', required=True, nargs='+', help='Path to nmap xml report.')
    parser.add_argument('-o', '--output', help='Path to xlsx output.')
    args = parser.parse_args()

    xml_reports = []
    for tmp_path in args.reports:
        if tmp_path.endswith('.xml') and os.path.isfile(tmp_path):
            xml_reports.append(tmp_path)
        elif os.path.isdir(tmp_path):
            for file_path in os.listdir(tmp_path):
                if file_path.endswith('.xml'):
                    xml_reports.append(os.path.join(tmp_path, file_path))
        else:
            parser.print_help()
            print(f'\n[!] "{tmp_path}" not a file or a directory.')
            exit()

    reports = []
    for report in xml_reports:
        try:
            parsed = NmapParser.parse_fromfile(report)
        except NmapParserException as e:
            parsed = NmapParser.parse_fromfile(report, incomplete=True)

        setattr(parsed, 'source', os.path.basename(report))
        reports.append(parsed)

    xlsx_path = args.output if args.output else f'Report_%s' % datetime.now().strftime('%Y%m%d_%H%M%S')
    if not xlsx_path.endswith('.xlsx'):
        xlsx_path += '.xlsx'
    workbook = Workbook(xlsx_path)
    main(reports, workbook)
