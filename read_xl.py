#!/usr/bin/env python
"""
Crudely:
1. Read Excel data and column types.
2. Dump a Django models.py file.

Requires: xlrd.
"""
from xlrd import open_workbook, XL_CELL_EMPTY, XL_CELL_TEXT, XL_CELL_NUMBER, \
    XL_CELL_DATE, XL_CELL_BOOLEAN, XL_CELL_ERROR, XL_CELL_BLANK

TYPE_LOOKUP = {
    XL_CELL_EMPTY: 'Empty',
    XL_CELL_TEXT: 'Text',
    XL_CELL_NUMBER: 'Number',
    XL_CELL_DATE: 'Date',
    XL_CELL_BOOLEAN: 'Boolean',
    XL_CELL_ERROR: 'Error',
    XL_CELL_BLANK: 'Blank',
    }

AMBIGUOUS_FIELD = 'CharField'
DJANGO_TYPE_LOOKUP = {
    XL_CELL_TEXT: 'CharField',
    XL_CELL_NUMBER: 'FloatField',
    XL_CELL_DATE: 'DateTimeField',
    XL_CELL_BOOLEAN: 'BooleanField',
    }

def dump_workbook(filename):
    wb = open_workbook(filename)
    for s in wb.sheets():
        print '-------------------------------------------------------------------'
        print 'Sheet:',s.name
        for row in range(s.nrows):
            values = []
            for col in range(s.ncols):
                values.append('%s' % s.cell(row,col).value)
            print ','.join(values)
        print

def dump_column_headers(filename):
    wb = open_workbook(filename)
    for s in wb.sheets():
        print '-------------------------------------------------------------------'
        print('Sheet:',s.name)
        for col in range(s.ncols):
            name = s.cell(0, col)
            types = {} # for unique
            for row in range(1, s.nrows):
                atype = TYPE_LOOKUP[s.cell(row, col).ctype]
                types.setdefault(atype, True)
            typenames = ','.join(types.keys())
            print('Column %d: "%s" (%s)' % (col, name, typenames))
        print()

def get_safename(name):
    tmp = name.title()
    new = []
    for ch in tmp:
        if ch.isalnum():
            new.append(ch)
    return ''.join(new)

def safe_js(value):
    tmp = value
    tmp = tmp.replace('\\', '\\\\')
    tmp = tmp.replace('"', '\\"')
    return tmp

last_struct = None
def get_workbook_structure(filename):
    """
    Returns a dictionary containing the workbook structure like this:
    { sheetname: [
        (column_header, types),
        ]
    }
    """
    global last_struct
    if last_struct is not None:
        return last_struct
    struct = {}
    wb = open_workbook(filename)
    totalrows = 0
    for s in wb.sheets():
        totalrows += s.nrows
        sheetname = s.name
        print 'Found Sheet "{0}"'.format(sheetname)
        struct.setdefault(sheetname, [])
        for col in range(s.ncols):
            name = s.cell(0, col).value
            print '   Column "{0}": '.format(name),
            types = []
            for row in range(1, s.nrows):
                atype = DJANGO_TYPE_LOOKUP.get(s.cell(row, col).ctype, None)
                if atype is not None:
                    types.append(atype)
            types = list(set(types)) # unique
            print ', '.join(types)
            struct[sheetname].append( (name, types) )
    print '{0} Totals Rows'.format(totalrows)
    last_struct = struct
    return struct

def dump_django_models(filename, outfilename):
    out = open(outfilename, "w")
    out.write('"""\n')
    out.write('Django models auto-generated from Excel spreadsheet:\n')
    out.write('{0}\n'.format(filename))
    out.write('"""\n')
    out.write('from django.db import models\n')
    struct = get_workbook_structure(filename)
    sheetnames = struct.keys()
    sheetnames.sort()
    for sheetname in sheetnames:
        col_defs = struct[sheetname]
        out.write('\n\n')
        modelname = get_safename(sheetname)
        out.write('class %s(models.Model):\n' % modelname )
        for name, types in col_defs:
            hashmark = ''
            alt = ''
            default_params = 'null=True, blank=True'
            safename = get_safename(name)
            if len(types) > 0 and AMBIGUOUS_FIELD != types[0]:
                # Make sure than ambiguous types get created as textfields.
                if AMBIGUOUS_FIELD in types:
                    types.remove(AMBIGUOUS_FIELD) # wherever it may be
                types.insert(0, AMBIGUOUS_FIELD) # at start of list
            for typename in types:
                params = default_params
                if typename in ['CharField']:
                    params += ', max_length=255'
                out.write('    %s%s = models.%s(%s) %s\n' % (hashmark, safename, typename, params, alt))
                alt = '# Alternate field type.'
                hashmark = '#'
    out.close()

def dump_django_admin(filename, outfilename):
    out = open(outfilename, "w")
    out.write('"""\n')
    out.write('Django admin auto-generated from Excel spreadsheet:\n')
    out.write(filename)
    out.write('\n"""\n')
    out.write('from django.contrib import admin\n')
    struct = get_workbook_structure(filename)
    sheetnames = struct.keys()
    modelnames = []
    fieldnames = {} #modelname: list of fields
    for sheetname in sheetnames:
        modelname = get_safename(sheetname)
        modelnames.append(modelname)
        fieldnames[modelname] = [
            get_safename(name) for name, types in struct[sheetname]
            if types
            ]
    modelnames.sort()
    out.write( 'from models import %s\n' % (', '.join(modelnames) ))
    out.write('\n')

    for modelname in modelnames:
        out.write( 'class {0}Admin(admin.ModelAdmin):\n'.format(modelname))
        fieldlist = ', '.join( "'{0}'".format(name) for name in fieldnames[modelname] )
        out.write( "    list_display = ('pk', {0})\n".format(fieldlist))
        out.write('\n\n')

    for modelname in modelnames:
        out.write('admin.site.register({0}, {0}Admin)\n'.format(modelname) )
    out.close()


def dump_json_fixture(filename, outfilename, appname):
    out = open(outfilename, "w")

# Django can't deserialize javascript comments.
#    out.write('/*')
#    out.write('* JSON fixture auto-generated from Excel spreadsheet:')
#    out.write('* ' + filename)
#    out.write('*/')
    out.write('[\n')
    struct = get_workbook_structure(filename)
    wb = open_workbook(filename)
    sheetnames = struct.keys()
    sheetnames.sort()
    more = False
    for sheetname in sheetnames:
        modelname = get_safename(sheetname)
        col_defs = struct[sheetname]
        colnames = [ get_safename(x) for x, y in col_defs ]
        s = wb.sheet_by_name(sheetname)

        # Calculate columns with data.
        lastcol = None
        columns_with_data = []
        for i in range(0, len(col_defs)):
            colname, types = col_defs[i]
            if len(types) < 0:
                # It's an empty column.
                continue
            columns_with_data.append(i)
            lastcol = i

        lastrow = s.nrows
        for row in range(1, s.nrows):
            if more:
                out.write( '  },\n')
            out.write( '  {\n')
            out.write( '    "pk": {0},\n'.format(row))
            out.write( '    "model": "{0}.{1}",\n'.format(appname, modelname.lower()))
            out.write( '    "fields": {\n')
            output = []
            for col in columns_with_data:
                try:
                    datum = '%s' % s.cell(row, col).value
                    datum = safe_js(datum)
                except:
                    continue
                colname = colnames[col]
                if datum == '' or colname == '':
                    continue
                fmt = '      "{0}": "{1}"'
                output.append(fmt.format(colname, datum))
            out.write( ',\n'.join(output))
            out.write( '    }\n')
            more = True
    if more:
        out.write( '  }\n')
    out.write(']\n')
    out.close()

def main():
    from optparse import OptionParser
    parser = OptionParser()
    options, args = parser.parse_args()
    try:
        inputfilename = args[0]
        outputappname = args[1]
        print 'Excel Input File: ' + inputfilename
        print 'Output App Name : ' + outputappname
    except:
        print 'Usage:'
        print '  read_xl.py INPUTFILENAME OUTPUTAPPNAME'

    models = 'output/{0}/models.py'.format(outputappname)
    admin = 'output/{0}/admin.py'.format(outputappname)
    fixture = 'output/{0}/converted.json'.format(outputappname)

    dump_django_models(inputfilename, models)
    print 'Created Django Models : ' + models
    dump_django_admin(inputfilename, admin)
    print 'Created Django Admin  : ' + admin
    dump_json_fixture(inputfilename, fixture, outputappname)
    print 'Dumped data to fixture: ' + fixture
    print 'DONE. Run syncdb and loaddata now.'

def debug():
    x = 'input/Tracking.xls'
    dump_workbook(x)
    dump_column_headers(x)


if __name__=='__main__':
    main()
