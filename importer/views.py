# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from django.shortcuts import render
from . import utils

import xlrd
import os


# Create your views here.

def importer_home(request):
    ctx = {}
    if request.method == 'POST':
        try:

            utils.country_finder('43212')

            if 'my_excel in request.FILES':
                ctx['results'] = 'True'
                input_file = request.FILES['my_excel']

                workbook = xlrd.open_workbook(filename=None, file_contents=input_file.read())
                sheet = workbook.sheet_by_index(0)
                template = open(os.path.join(os.path.dirname(os.path.dirname(__file__)), 'templates/importer/partials/tables.html'), 'w')

                for x in range(1, sheet.nrows-1):

                    template.write('<tr>\n')

                    template.write('<td>\n')
                    template.write(str(x))
                    template.write('</td>\n')

                    template.write('<td>')
                    template.write(str(sheet.col(2)[x].value).partition('.')[0])
                    template.write(" - ")
                    template.write(str(sheet.col(1)[x].value).partition('.')[0])
                    template.write('</td>\n')

                    template.write('<td>')
                    template.write(str(sheet.col(4)[x].value).partition('.')[0])
                    template.write("-")
                    template.write(str(sheet.col(3)[x].value).partition('.')[0])
                    template.write('</td>\n')

                    template.write('<td>')
                    template.write(str(sheet.col(5)[x].value).partition('.')[0])
                    template.write('</td>\n')

                    template.write('<td>')
                    template.write(str(sheet.col(6)[x].value).partition('.')[0])
                    template.write('</td>\n')


                    template.write('</tr>\n')

                template.close()

        except IndexError:
            ctx['results'] = 'False'
            ctx['message'] = 'Wrong File! File does not have enough values to unpack. Please ensure file has right ' \
                             'values and try again!'
    else:
        pass

    return (render(request, 'importer/index.html', ctx))