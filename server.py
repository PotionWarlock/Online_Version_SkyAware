from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import os

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "Insight.xlsx")
if not os.path.exists(EXCEL_PATH):
    raise FileNotFoundError(f"{EXCEL_PATH} not found.")

# Load both sheets
df_summary = pd.read_excel(EXCEL_PATH, sheet_name="Summary")
df_history = pd.read_excel(EXCEL_PATH, sheet_name="HistoryData")

# Convert Date column to datetime if it exists
if 'Date' in df_history.columns:
    df_history['Date'] = pd.to_datetime(df_history['Date'])

@app.route('/')
def serve_html():
    return send_from_directory('.', 'index.html')

@app.route('/api/states', methods=['GET'])
def get_states():
    try:
        return jsonify(df_summary['State'].tolist())
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/api/state', methods=['POST'])
def receive_state():
    data = request.get_json(force=True)
    state = data.get('state', None)

    if not state:
        return jsonify({"error": "No state provided."})

    row = df_summary[df_summary['State'] == state]
    if row.empty:
        return jsonify({"error": f"No data found for {state}."})

    surface_pressure = float(row['Surface Pressure'].values[0])
    max_formaldehyde = float(row['MAX Formaldehyde (molecules/cm^2)'].values[0])
    min_formaldehyde = float(row['MIN Formaldehyde (molecules/cm^2)'].values[0])

    return jsonify({
        "surface_pressure": surface_pressure,
        "max_formaldehyde": max_formaldehyde,
        "min_formaldehyde": min_formaldehyde
    })

@app.route('/api/history/<state>', methods=['GET'])
def get_history(state):
    try:
        # Filter historical data for the selected state
        state_history = df_history[df_history['State'] == state]
        
        if state_history.empty:
            return jsonify({"error": f"No historical data found for {state}."})
        
        # Sort by date
        state_history = state_history.sort_values('Date')
        
        # Prepare the data for JSON response
        history_data = {
            "dates": state_history['Date'].dt.strftime('%Y-%m-%d').tolist(),
            "surface_pressure": state_history['Surface Pressure'].fillna(0).tolist(),
            "max_formaldehyde": state_history['MAX Formaldehyde (molecules/cm^2)'].fillna(0).tolist(),
            "min_formaldehyde": state_history['MIN Formaldehyde (molecules/cm^2)'].fillna(0).tolist()
        }
        
        # Calculate trends
        trends = calculate_trends(state_history)
        
        return jsonify({
            "history": history_data,
            "trends": trends
        })
        
    except Exception as e:
        return jsonify({"error": str(e)})

def calculate_trends(state_history):
    """Calculate trends from historical data"""
    if len(state_history) < 2:
        return {"error": "Insufficient data for trend analysis"}
    
    trends = {}
    
    # Calculate trends for each metric
    metrics = ['Surface Pressure', 'MAX Formaldehyde (molecules/cm^2)', 'MIN Formaldehyde (molecules/cm^2)']
    
    for metric in metrics:
        if metric in state_history.columns:
            values = state_history[metric].dropna()
            if len(values) > 1:
                # Simple linear trend (last value - first value)
                first_val = values.iloc[0]
                last_val = values.iloc[-1]
                change = last_val - first_val
                change_percent = (change / first_val * 100) if first_val != 0 else 0
                
                trends[metric] = {
                    "first_value": float(first_val),
                    "last_value": float(last_val),
                    "change": float(change),
                    "change_percent": float(change_percent),
                    "trend_direction": "increasing" if change > 0 else "decreasing" if change < 0 else "stable"
                }
    
    # Overall air quality trend
    if 'MAX Formaldehyde (molecules/cm^2)' in trends:
        formaldehyde_trend = trends['MAX Formaldehyde (molecules/cm^2)']
        trends['air_quality_trend'] = {
            "direction": formaldehyde_trend['trend_direction'],
            "description": f"Formaldehyde levels are {formaldehyde_trend['trend_direction']} by {abs(formaldehyde_trend['change_percent']):.1f}%"
        }
    
    return trends

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)  # Changed for Render