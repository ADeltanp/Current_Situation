from visdom import Visdom
import numpy as np
import openpyxl as opx

from plotly.subplots import make_subplots
import plotly.graph_objects as go


def read_xls(xls_path):
    xls = opx.load_workbook(xls_path)
    sheet = xls[xls.sheetnames[0]]
    imports = list()
    exports = list()
    for row in sheet.rows:
        if row[2].value != 'China':
            continue
        imports.append([cell.value for cell in row[3:16]])
        exports.append([cell.value for cell in row[16:]])
    imports = np.stack(imports)
    exports = np.stack(exports)
    return imports, exports


def visualize(imports, exports, vis):
    year_imports = imports[:, 12]
    year_exports = exports[:, 12]
    wo1112_imports = imports[:, :10].sum(axis=1)
    wo1112_exports = exports[:, :10].sum(axis=1)

    year_deficit = year_imports - year_exports
    year_line = np.stack((year_imports, year_exports, year_deficit), axis=1)
    wo1112_deficit = wo1112_imports - wo1112_exports
    wo1112_line = np.stack((wo1112_imports, wo1112_exports, wo1112_deficit), axis=1)

    x = np.arange(1985, 2020)
    strip_year_line = year_line[:-1]
    strip_x = np.arange(1985, 2019)
    print('2018 deficit minus 2017 deficit = ', year_deficit[-2] - year_deficit[-3])

    vis.line(strip_year_line, strip_x, win='imp & exp w\\o 2019',
             opts=dict(
                 legend=['import', 'export', 'deficit'],
                 xlabel='Year',
                 ylabel='Million USD'
             ))

    vis.line(wo1112_line, x,
             win='w\\o 11 12',
             opts=dict(
                 legend=['import', 'export', 'deficit'],
                 xlabel='Year',
                 ylabel='Million USD'
             ))


def monthly_plot(imports, exports, vis):
    deficits = imports - exports
    fig = make_subplots(rows=3, cols=4,
                        subplot_titles=('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                                        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'))
    x = np.arange(1985, 2020)
    for j in range(10):
        row = int(j / 4) + 1
        col = int(j % 4) + 1
        fig.add_trace(go.Scatter(x=x, y=imports[:, j],
                                 legendgroup='group1',
                                 name='import',
                                 mode='lines',
                                 line=dict(color='#1F77B4'),
                                 showlegend=j == 0),
                      row=row,
                      col=col)
        fig.add_trace(go.Scatter(x=x, y=exports[:, j],
                                 legendgroup='group2',
                                 name='export',
                                 mode='lines',
                                 line=dict(color='#FF7F0E'),
                                 showlegend=j == 0),
                      row=row,
                      col=col)
        fig.add_trace(go.Scatter(x=x, y=deficits[:, j],
                                 legendgroup='group3',
                                 name='deficit',
                                 mode='lines',
                                 line=dict(color='#2CA02C'),
                                 showlegend=j == 0),
                      row=row,
                      col=col)

    for j in range(2):
        fig.add_trace(go.Scatter(x=x,
                                 y=imports[:-1, 10 + j],
                                 legendgroup='group1',
                                 name='import',
                                 mode='lines',
                                 line=dict(color='#1F77B4'),
                                 showlegend=False),
                      row=3,
                      col=3 + j)
        fig.add_trace(go.Scatter(x=x,
                                 y=exports[:-1, 10 + j],
                                 legendgroup='group2',
                                 name='export',
                                 mode='lines',
                                 line=dict(color='#FF7F0E'),
                                 showlegend=False),
                      row=3,
                      col=3 + j)
        fig.add_trace(go.Scatter(x=x,
                                 y=deficits[:-1, 10 + j],
                                 legendgroup='group3',
                                 name='deficit',
                                 mode='lines',
                                 line=dict(color='#2CA02C'),
                                 showlegend=False),
                      row=3,
                      col=3 + j)

    vis.plotlyplot(fig)


viz = Visdom(env='Trade')
i, e = read_xls('E:\\Course\\XS\\data\\country.xlsx')
visualize(i, e, viz)
monthly_plot(i, e, viz)
