'''
The pypsa_academy toolbox is a collection of functions designed to enhance the usability
of Python for Power System Analysis (PyPSA). These functions are specifically developed 
to simplify PyPSA's application, making it more accessible and user-friendly, particularly 
for teaching and educational purposes.

For any questions, contact: 
Priyesh Gosai
priyesh@innovateimpact.com
 
Version 2.0.0 - 17 October 2024
'''

import os
import pandas as pd
import numpy as np
import pypsa
from ipywidgets import interact, fixed
from IPython.display import display, clear_output
import ipywidgets as widgets
import logging
from concurrent.futures import ThreadPoolExecutor
import plotly.graph_objects as go
import warnings
import requests


def convert_sheet_to_csv(xls, sheet_name, csv_folder_path):
    df = xls.parse(sheet_name)
    csv_file_path = os.path.join(csv_folder_path, f"{sheet_name}.csv")
    df.to_csv(csv_file_path, index=False)
    logging.info(f"Converted {sheet_name} to CSV.")
    return csv_file_path

def convert_excel_to_csv(excel_file_path, csv_folder_path):
    """
    Converts each sheet in an Excel file to a CSV file, only for sheets whose names are in a predefined list. 
    The function checks if the target folder exists, and only specific CSV files related to the Excel file's 
    sheets are deleted and recreated.

    Parameters:
    excel_file_path (str): The file path of the Excel file.
    csv_folder_path (str): The path to the folder where CSV files will be saved.

    Returns:
    List[str]: Paths to the successfully created CSV files.
    """

    logging.basicConfig(level=logging.INFO)
    # components = {"buses", "carriers", "generators", "generators-p_max_pu",
    # "generators-p_min_pu", "generators-p_set","line_types", "lines",
    # "links", "links-p_max_pu","links-p_min_pu", "links-p_set",
    # "loads", "loads-p_set", "shapes","shunt_impedances",
    # "snapshots", "storage_units", "stores","sub_networks",
    # "transformer_types","transformers"}

    import pypsa
    n = pypsa.Network()
    components = {
        f"{n.components[key]['list_name']}-{dict_value}" if dict_value else n.components[key]['list_name']
        for key in n.component_attrs
        for dict_value in list(n.component_attrs[key][(n.component_attrs[key]['status'].str.contains("Input")) &\
                                                       (n.component_attrs[key]['varying'] == True)].index) or [None]
                }
    
    created_csv_files = []

    # Ensure the CSV folder exists
    os.makedirs(csv_folder_path, exist_ok=True)

    # Clear only relevant CSV files in the folder
    for item in os.listdir(csv_folder_path):
        if item.endswith(".csv") and item.replace(".csv", "") in components:
            os.remove(os.path.join(csv_folder_path, item))

    try:
        xls = pd.ExcelFile(excel_file_path)
        with ThreadPoolExecutor() as executor:
            futures = [executor.submit(convert_sheet_to_csv, xls, sheet_name, csv_folder_path)
                       for sheet_name in xls.sheet_names if sheet_name in components]
            for future in futures:
                created_csv_files.append(future.result())
    except Exception as e:
        logging.error(f"Error converting Excel to CSV: {e}")
        return []
    finally:
        if xls is not None:
            xls.close()
            print('Excel file is closed')

    logging.info(f"Conversion complete. CSV files are saved in '{csv_folder_path}'")
    return csv_folder_path

def solver_selected():

    def simple_network(solver_name):
        try:
            # Initialize the test_network
            test_network = pypsa.Network()

            # Add snapshots (3 hours)
            snapshots = pd.date_range('2024-01-01 00:00', periods=3, freq='h')
            test_network.set_snapshots(snapshots)

            test_network.add("Carrier", "AC")

            # Add buses
            test_network.add("Bus", "Bus1", carrier='AC')
            test_network.add("Bus", "Bus2", carrier='AC')
            test_network.add("Bus", "Bus3", carrier='AC')

            # Add links between buses
            test_network.add("Link", "Link1", bus0="Bus1", bus1="Bus2", p_nom=100)
            test_network.add("Link", "Link2", bus0="Bus2", bus1="Bus3", p_nom=100)

            # Add load on Bus3
            test_network.add("Load", "Load1", bus="Bus3", p_set=[10, 15, 20])

            # Add generators
            test_network.add("Generator", "Gen1", bus="Bus1", p_nom=30, marginal_cost=50)
            test_network.add("Generator", "Gen2", bus="Bus2", p_nom=40, marginal_cost=30)

            # Solve the test_network
            solved = test_network.optimize(solver_name=solver_name)
            
            # Clear output after solving (optional for cleaner output)
            clear_output(wait=True)
            
            return solved  # Return the solved network object if successful
        except Exception as e:
            print(f"Solver {solver_name} failed: {e}")
            return None  # Return None if the solver fails


    def find_solver():
        # List solvers to test, ordered by preference
        solver_options = ['gurobi', 'cplex','mosek', 'highs','glpk']

        # Iterate over the solver options and try to solve the simple network
        for solver in solver_options:
            print(f"Testing solver: {solver}")
            result = simple_network(solver_name=solver)
            if result is not None:
                print(f"Solver {solver} succeeded!")
                return solver  # Return the first working solver
        
        # If no solver works, raise an error
        raise ValueError("No suitable solver found. Please install one of the solvers: " + ', '.join(solver_options))


    # Example usage
    try:
        # Find the first working solver by testing them all
        solver = find_solver()
        print(f"Selected solver: {solver}")
    except ValueError as e:
        print(e)
    return solver

def pypsa_viewer(network):
    # dropdown menu: Component List
    select_component = list(network.all_components)
    select_component.sort()
    io_types = ['Input','Output']
    type_options = ['static','varying']
    view_type = ['Plot','Table']

    clear_output()

    # Define the initial dropdown and radio buttons
    dropdown = widgets.Dropdown(
        options=select_component,
        value=select_component[0],
        description='Component:',
        disabled=False,
    )

    radio1 = widgets.RadioButtons(
        options=io_types,
        description=io_types[0],
        disabled=False
    )

    radio2 = widgets.RadioButtons(
        options=type_options,
        description=type_options[0],
        disabled=False
    )

    # Initialize the shared dictionary to hold the return file
    state = {"return_file": None}

    # Define a function to handle the user's confirmation
    def on_confirm_clicked(b):
        val = dropdown.value
        io_data = radio1.value
        type_var = radio2.value

        varying_attributes = list(network.component_attrs[val][(network.component_attrs[val]['status'].str.contains(io_data)) & network.component_attrs[val]['varying'] == True].index)
        static_attributes = list(network.component_attrs[val][network.component_attrs[val]['status'].str.contains(io_data)].index)
        if 'name' in static_attributes:
            static_attributes.remove('name')

        # Conditional logic for second radio button
        if type_var == type_options[1]:
            # Create additional dropdown and radio button widgets
            new_dropdown = widgets.Dropdown(
                options=varying_attributes,
                value=varying_attributes[0],
                description='Variables:',
                disabled=False,
            )
            new_radio = widgets.RadioButtons(
                options=view_type,
                value=view_type[0],
                description="View",
                disabled=False
            )

            # Function to handle the new widget interaction
            def on_new_confirm_clicked(new_b):
                new_selected_dropdown = new_dropdown.value
                new_selected_radio = new_radio.value
                clear_output()

                if type_var == 'varying':
                    time_data = getattr(network, f"{network.components[val]['list_name']}_t")[new_selected_dropdown]
                    if new_selected_radio == 'Table':
                        state['return_file'] = time_data
                        display(time_data)
                    elif new_selected_radio == 'Plot':
                        if time_data.empty:
                            print('No data to plot')
                            state['return_file'] = None
                        else:
                            # Plot using Plotly explicitly with traces for each column
                            fig = go.Figure()
                            for col in time_data.columns:
                                fig.add_trace(go.Scatter(x=time_data.index, y=time_data[col], mode='lines', name=col))
                            fig.update_layout(title=f"{val} {new_selected_dropdown} Time Series", xaxis_title="Time", yaxis_title="Value")
                            fig.show()
                            state['return_file'] = time_data


            # Button to confirm the new selections
            new_confirm_button = widgets.Button(description="Confirm New Selections")
            new_confirm_button.on_click(on_new_confirm_clicked)

            # Display the new widgets
            display(new_dropdown, new_radio, new_confirm_button)

        else:
            clear_output()
            component_data = getattr(network, network.components[val]['list_name'])

            # Filter the static attributes to only include those that exist in the columns
            valid_attributes = [attr for attr in static_attributes if attr in network.generators.columns]

            # Display only the valid attributes
            if valid_attributes:
                state['return_file'] = component_data[valid_attributes]
                display(component_data[valid_attributes])
            else:
                print("Warning: None of the static attributes are found in network.generators columns")

    # Button to confirm the initial selections
    confirm_button = widgets.Button(description="Confirm Selections")
    confirm_button.on_click(on_confirm_clicked)

    # Display the initial widgets
    display(dropdown, radio1, radio2, confirm_button)

    # Return state for further use if needed
    return state

