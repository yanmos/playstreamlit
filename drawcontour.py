#!/usr/bin/python3

import streamlit as st
import numpy as np
import openpyxl as px
import matplotlib.pyplot as plt
from pydantic import NoneIsAllowedError, NoneIsNotAllowedError
from scipy import interpolate
from io import BytesIO
import tempfile
from openpyxl.drawing.image import Image

HEADER_POS = 'C5'
GRAPH_POS = 'M6'
GRAPH_POS_DEF = 26
N_POINTS = 2000
N_LEVEL = 11     # N + 1
G_INIT = 0
G_UPLOADED = 1
G_END = 10

def main(state):
    st.header("Draw contour into .xlsx")
    tempfilename = "drawcontour_template.xlsx"
    with open(tempfilename, "rb") as fdata:
        st.download_button("Download template .xlsx file",
                        data=fdata,
                        mime='xlsx',
                        file_name=tempfilename)
    st.session_state['state'] = G_INIT
    upfile = st.file_uploader("Upload xlsx file",
                type="xlsx",
                accept_multiple_files=False,
                key=None,
                help=None,
                on_change=None,
                args=None,
                kwargs=None,
                disabled=False)
    if upfile:
        st.session_state['state'] = G_UPLOADED
        with tempfile.NamedTemporaryFile() as tmp:
            st.session_state['state'] = draw_contour(upfile, tmp.name)
            fdata = BytesIO(tmp.read())
        if st.session_state['state'] == G_END:
            st.download_button("Download Result file",
                            data=fdata,
                            mime='xlsx',
                            file_name=upfile.name)

def draw_contour(upfile, outfilename):
    wb = px.load_workbook(upfile, data_only=True)
    sheetnames = wb.sheetnames
    sheetlist = {}
    res = st.session_state['state']
    for sheetname in sheetnames:
        sheetlist[sheetname] = st.checkbox(sheetname,
                                value=False, 
                                key=sheetname, 
                                help=None, 
                                on_change=None, 
                                args=None, 
                                kwargs=None, 
                                disabled=False)
    if st.button('Draw Contour'):
        st.write('Drawing...')
        for sheetname in sheetnames:
            if sheetlist[sheetname]:
                contour(upfile, sheetname, outfilename)
        res = G_END
    return res

def contour(fileobj, in_shname, outfilename):
    wb, ws, x_label, y_label, z_label, x, y, z, cmaps, methods = set_parameter(fileobj, in_shname)
    make_contour(wb, ws, x_label, y_label, z_label, x, y, z, cmaps, methods)
    wb.save(outfilename)

def set_parameter(in_fpath, in_shname):
    wb = px.load_workbook(in_fpath, data_only=True)
    ws = wb[in_shname]
    wb.active = ws
    x_label = ws[HEADER_POS].value
    y_label = ws[HEADER_POS].offset(0, 1).value
    z_label = []
    methods = []
    cmaps = []
    x = np.array([])
    y = np.array([])
    for col in ws.iter_cols(min_row=5, min_col=5, max_row=5):
        for head in col:
            if not head.value is None:
                z_label.append(head.value)
                methods.append(head.offset(-2, 0).value)
                cmaps.append(head.offset(-3, 0).value)
    for row in ws.iter_rows(min_row=6, min_col=3, max_col=4):
        if is_num(row[0].value):
            x = np.r_[x, float(row[0].value)]
        if is_num(row[1].value):
            y = np.r_[y, float(row[1].value)]
    n_z = len(z_label)
    n_data = np.size(x, 0)
    z = np.array([])
    for col in ws.iter_cols(min_row=6, min_col=5, max_row=5+n_data, max_col=4+n_z):
        if not col[0].value is None:
            zi = np.array([])
            if np.size(z) == 0:
                for data in col:
                    if is_num(data.value):
                        zi = np.r_[zi, float(data.value)]
                z = np.array([zi,])
            else:
                for data in col:
                    if is_num(data.value):
                        zi = np.r_[zi, float(data.value)]
                z = np.r_[z, [zi]]
    wb = px.load_workbook(in_fpath, data_only=False)
    ws = wb[in_shname]
    wb.active = ws
    return wb, ws, x_label, y_label, z_label, x, y, z, cmaps, methods

def make_contour(wb, ws, x_label, y_label, z_label, x, y, z, cmaps, methods):
    wGraph = ws[GRAPH_POS]
    pos = wGraph
    xmin = 0    # np.min(x)
    xmax = np.max(x)
    ymin = 0    # np.min(y)
    ymax = np.max(y)
    x1 = np.linspace(xmin, xmax + xmax/N_POINTS, N_POINTS)
    y1 = np.linspace(ymin, ymax + ymax/N_POINTS, N_POINTS)
    xx, yy = np.meshgrid(x1, y1)
    zz = np.array([])
    zzi = np.array([])
    n_z = len(z_label)
    for i in range(n_z):
        if np.size(zz) == 0:
            zz = interpolate.griddata((x,y), z[i], (xx, yy), method=methods[i])
            zz = np.stack([zz,])
        else:
            zzi = interpolate.griddata((x,y), z[i], (xx, yy), method=methods[i])
            zz = np.concatenate([zz, [zzi]])
    for i in range(n_z):
        zmax = np.max(z[i])
        zmin = np.min(z[i])
        if zmax < 0:
            lv0 = zmin
            lv0 = -np.ceil(-zmin/10.0)*10.0
            lvf = 0
        elif zmin > 0:
            lv0 = 0
            lvf = zmax
            lvf = np.ceil(zmax/10.0)*10.0
        else:
            lv0 = zmin
            lv0 = -np.ceil(-zmin/10.0)*10.0
            lvf = zmax
            lvf = np.ceil(zmax/10.0)*10.0
        fig = plt.figure()
        plt.contourf(xx, yy, zz[i], levels=np.linspace(lv0, lvf, N_LEVEL), cmap=cmaps[i])
        plt.title(z_label[i])
        plt.xlabel(x_label)
        plt.ylabel(y_label)
        plt.grid(which='major', color='gray', linestyle='-')
        plt.colorbar()
        buffer = BytesIO()
        fig.savefig(buffer, format='png')
        buffer.seek(0)
        img = Image(buffer)
        ws.add_image(img, pos.coordinate)
        pos = pos.offset(GRAPH_POS_DEF, 0)
        st.write(fig)

def is_num(s):
    try:
        float(s)
    except ValueError:
        return False
    except NoneIsNotAllowedError:
        return False
    except NoneIsAllowedError:
        return False
    except TypeError:
        return False
    else:
        return True

if __name__ == '__main__':
    main(G_INIT)
