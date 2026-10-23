"""
bokeh_read_and_write_file.py: file input and output with Bokeh server
"""

# run in server mode via: bokeh serve --show bokeh_read_and_write_file.py
# then click FileInput widget and choose a csv file with two columns

from bokeh.io import curdoc
from bokeh.layouts import column, row
from bokeh.models.widgets import FileInput, Button
from bokeh.models import ColumnDataSource, CustomJS
from bokeh.plotting import figure
from base64 import b64decode
import pandas as pd
import numpy as np
import io

def plot_data():
    plot.line('x', 'y', source=datasource)

def upload_data(attr, old, new):
    global datasource, plot
    decoded = b64decode(new)
    data = io.BytesIO(decoded)
    df = pd.read_csv(data, header=None)
    df.columns = ('x', 'y')
    datasource.data = {'x': df['x'], 'y': df['y']}
    plot_data()

def subtract_mean_background():
    global datasource
    datasource.data['y'] = datasource.data['y'] - np.mean(datasource.data['y'])
    plot_data()

datasource = ColumnDataSource({'x': [], 'y': []})
plot = figure(width=500, height=400)

file_input = FileInput(accept=".csv")
file_input.on_change('value', upload_data)

background_subtract_button = Button(label="Subtract background", button_type="primary", width=150)
background_subtract_button.on_click(subtract_mean_background)

jscode = """
var data = source.data;
var x = data['x'];
var y = data['y'];
console.log(x);
console.log(y);
// comment
var filetext = "";
var lines = [];
for (var i = 0; i < x.length; i++) {
    var currRow = x[i].toString() + "," + y[i].toString();
    lines.push(currRow);
}
filetext = lines.join('\\n');

const filename = 'data_result.csv';
var blob = new Blob([filetext], { type: 'text/csv;charset=utf-8;' });

//addresses IE
if (navigator.msSaveBlob) {
    navigator.msSaveBlob(blob, filename);
}

else {
    var link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.style.visibility = 'hidden';
    link.dispatchEvent(new MouseEvent('click'));
}
"""

download_button = Button(label="Download processed data", button_type="primary", width=200)
download_button.js_on_click(CustomJS(args=dict(source=datasource),
                                         code=jscode))

layout = column(file_input, row(background_subtract_button, download_button), plot)

curdoc().add_root(layout)