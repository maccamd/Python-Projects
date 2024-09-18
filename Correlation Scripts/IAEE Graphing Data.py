from dash import Dash, html, dcc, callback, Output, Input
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd

df = pd.read_excel('TROC Emission Test Tracker.xlsm')
df = df.drop(columns=['Comments', 'Barometer', 'Test Name', 'Date', 'Mileage', 'Violations','Temperature'])
vdc1df = df[(df.Chamber == 'VDC1') & (df.Conformity =='Y')]
vdc2df = df[(df.Chamber == 'VDC2') & (df.Conformity =='Y')]
#print(vdc2df.head())

app = Dash()

fig = make_subplots(rows=2, cols=1,shared_xaxes=True, subplot_titles=("CO mg/Km All Data Points", "NOx mg/Km All Data Points"))

fig.append_trace(go.Scatter(x=vdc1df.Index, y=vdc1df.CO_km, name="CO/km VDC1"),1,1)
fig.append_trace(go.Scatter(x=vdc2df.Index, y=vdc2df.CO_km, name="CO/km VDC2"),1,1)

fig.append_trace(go.Scatter(x=vdc1df.Index, y=vdc1df.Nox_km, name="NOx/km VDC1"),2,1)
fig.append_trace(go.Scatter(x=vdc2df.Index, y=vdc2df.Nox_km, name="NOx/km VDC2"),2,1)

fig.show()

if __name__ == '__main__':
    app.run(debug=True)
