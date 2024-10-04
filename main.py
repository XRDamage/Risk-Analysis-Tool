import customtkinter as ctk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import filedialog
import numpy as np

# Initialize risk data list
risk_data = []
overall_risk_score = 0
max_possible_risk = 5 * 5  # Maximum possible risk score (Impact 5 * Likelihood 5)

# Function to calculate risk score
def calculate_risk(impact, likelihood):
    return impact * likelihood

# Function to calculate the overall risk as a percentage
def calculate_risk_percentage(total_risk_score, max_total_risk_score):
    return (total_risk_score / max_total_risk_score) * 100

# Function to load threats from Excel with enhanced error handling
def load_threats_from_excel():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
    
    if not file_path:
        label_message.configure(text="Error: No file selected.", fg_color="red")
        return

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        label_message.configure(text=f"Error: {str(e)}", fg_color="red")
        return

    required_columns = ["Threat Description", "Impact (1-5)", "Likelihood (1-5)", "Financial Impact ($)", "Threat Event Frequency (1-5)"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        label_message.configure(text=f"Error: Missing columns: {', '.join(missing_columns)}.", fg_color="red")
        return

    global overall_risk_score
    overall_risk_score = 0  # Reset the overall risk score

    # Clear previous data
    risk_data.clear()
    total_possible_risk_score = 0

    for index, row in df.iterrows():
        description = row["Threat Description"]
        impact = row["Impact (1-5)"]
        likelihood = row["Likelihood (1-5)"]
        financial_impact = row["Financial Impact ($)"]
        threat_frequency = row["Threat Event Frequency (1-5)"]

        # Validate impact and likelihood
        if not (1 <= impact <= 5) or not (1 <= likelihood <= 5):
            label_message.configure(text=f"Error: Invalid values at row {index + 1}.", fg_color="red")
            continue

        # Calculate risk score
        risk_score = calculate_risk(impact, likelihood)
        overall_risk_score += risk_score

        # Add the threat details to the risk_data list
        risk_data.append({
            "Threat ID": index + 1,
            "Threat Description": description,
            "Impact (1-5)": impact,
            "Likelihood (1-5)": likelihood,
            "Threat Event Frequency (1-5)": threat_frequency,
            "Financial Impact ($)": financial_impact,
            "Risk Score": risk_score
        })

        total_possible_risk_score += max_possible_risk

    # Calculate risk percentage
    risk_percentage = calculate_risk_percentage(overall_risk_score, total_possible_risk_score)

    label_message.configure(text=f"Threats loaded successfully. Overall Risk: {risk_percentage:.2f}%")
    display_threats()

# Function to display the loaded threats
def display_threats():
    # Clear the frame first
    for widget in frame_threats.winfo_children():
        widget.destroy()

    if risk_data:
        headers = ["Threat ID", "Threat Description", "Impact (1-5)", "Likelihood (1-5)", "Frequency (1-5)", "Financial Impact ($)", "Risk Score"]
        for col_num, header in enumerate(headers):
            ctk.CTkLabel(frame_threats, text=header, font=("Arial", 10, "bold")).grid(row=0, column=col_num, padx=5, pady=2)
        
        for row_num, threat in enumerate(risk_data, start=1):
            ctk.CTkLabel(frame_threats, text=threat["Threat ID"]).grid(row=row_num, column=0, padx=5, pady=2)
            ctk.CTkLabel(frame_threats, text=threat["Threat Description"]).grid(row=row_num, column=1, padx=5, pady=2)
            ctk.CTkLabel(frame_threats, text=threat["Impact (1-5)"]).grid(row=row_num, column=2, padx=5, pady=2)
            ctk.CTkLabel(frame_threats, text=threat["Likelihood (1-5)"]).grid(row=row_num, column=3, padx=5, pady=2)
            ctk.CTkLabel(frame_threats, text=threat["Threat Event Frequency (1-5)"]).grid(row=row_num, column=4, padx=5, pady=2)
            ctk.CTkLabel(frame_threats, text=threat["Financial Impact ($)"]).grid(row=row_num, column=5, padx=5, pady=2)
            ctk.CTkLabel(frame_threats, text=threat["Risk Score"]).grid(row=row_num, column=6, padx=5, pady=2)
    else:
        ctk.CTkLabel(frame_threats, text="No threats loaded.", font=("Arial", 10)).pack(pady=10)

# Function to open the mitigation window
def open_mitigation_window():
    mitigation_window = ctk.CTkToplevel(app)
    mitigation_window.title("Apply Mitigation")

    ctk.CTkLabel(mitigation_window, text="Select a Threat to Mitigate (Threat ID):", font=("Arial", 14)).pack(pady=10,padx=10)

    # Dropdown menu to select the threat by ID
    threat_ids = [threat["Threat ID"] for threat in risk_data]
    selected_threat_id = ctk.StringVar(value=threat_ids[0])
    dropdown = ctk.CTkOptionMenu(mitigation_window, values=[str(id) for id in threat_ids], variable=selected_threat_id)
    dropdown.pack(pady=10)

    ctk.CTkLabel(mitigation_window, text="Enter New Likelihood (1-5):", font=("Arial", 12)).pack(pady=5)
    likelihood_entry = ctk.CTkEntry(mitigation_window)
    likelihood_entry.pack(pady=5)

    ctk.CTkLabel(mitigation_window, text="Enter New Frequency (1-5):", font=("Arial", 12)).pack(pady=5)
    frequency_entry = ctk.CTkEntry(mitigation_window)
    frequency_entry.pack(pady=5)

    # Function to apply mitigation and recalculate risk
    def apply_mitigation():
        selected_id = int(selected_threat_id.get())
        new_likelihood = int(likelihood_entry.get())
        new_frequency = int(frequency_entry.get())

        # Update the threat's likelihood and frequency
        for threat in risk_data:
            if threat["Threat ID"] == selected_id:
                threat["Likelihood (1-5)"] = new_likelihood
                threat["Threat Event Frequency (1-5)"] = new_frequency
                threat["Risk Score"] = calculate_risk(threat["Impact (1-5)"], new_likelihood)
        
        # Recalculate overall risk
        global overall_risk_score
        overall_risk_score = sum(threat["Risk Score"] for threat in risk_data)
        risk_percentage = calculate_risk_percentage(overall_risk_score, len(risk_data) * max_possible_risk)
        
        label_message.configure(text=f"Risk recalculated. Overall Risk: {risk_percentage:.2f}%")
        display_threats()
        mitigation_window.destroy()

    # Button to apply the mitigation
    apply_button = ctk.CTkButton(mitigation_window, text="Apply Mitigation", command=apply_mitigation)
    apply_button.pack(pady=20)

# Function to plot the comparison of likelihood and impact in a new window
def plot_likelihood_vs_impact():
    # Open a new window for the plot
    plot_window = ctk.CTkToplevel(app)
    plot_window.title("Likelihood vs Impact Plot")

    impacts = [threat["Impact (1-5)"] for threat in risk_data]
    likelihoods = [threat["Likelihood (1-5)"] for threat in risk_data]
    threat_ids = [threat["Threat ID"] for threat in risk_data]  # Get the Threat IDs

    mean_likelihood = np.mean(likelihoods)
    mean_impact = np.mean(impacts)
    
    std_likelihood = np.std(likelihoods)
    std_impact = np.std(impacts)
    
    radius = (std_likelihood + std_impact) / 2

    fig, ax = plt.subplots(figsize=(8, 6))

    ax.scatter(likelihoods, impacts, color="blue", edgecolor="black")
    ax.set_xlabel("Likelihood (1-5)")
    ax.set_ylabel("Impact (1-5)")
    ax.set_title("Likelihood vs Impact Comparison")
    ax.grid(True)
    ax.set_xlim(1, 5)
    ax.set_ylim(1, 5)

    circle = plt.Circle((mean_likelihood, mean_impact), radius, color='green', fill=False, linestyle='--')
    ax.add_artist(circle)

    for i, threat_id in enumerate(threat_ids):
        ax.text(likelihoods[i], impacts[i], threat_id, fontsize=9, ha='right')

    canvas = FigureCanvasTkAgg(fig, master=plot_window)
    canvas.get_tk_widget().pack(pady=10)
    canvas.draw()

# Main Application Window
app = ctk.CTk()
app.title("Risk Management Tool")

# Main Layout Frames
frame_controls = ctk.CTkFrame(app)
frame_controls.pack(padx=10, pady=10)

frame_threats = ctk.CTkFrame(app)
frame_threats.pack(padx=10, pady=10, fill="x")

# Controls
button_load_threats = ctk.CTkButton(frame_controls, text="Load Threats", command=load_threats_from_excel)
button_load_threats.pack(side="left", padx=10)

button_plot = ctk.CTkButton(frame_controls, text="Plot Likelihood vs Impact", command=plot_likelihood_vs_impact)
button_plot.pack(side="left", padx=10)

button_mitigate = ctk.CTkButton(frame_controls, text="Mitigate Threat", command=open_mitigation_window)
button_mitigate.pack(side="left", padx=10)

label_message = ctk.CTkLabel(app, text="", font=("Arial", 12))
label_message.pack(pady=10)

app.mainloop()
