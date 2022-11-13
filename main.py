import pandas as pd
import os
from matplotlib import pyplot as plt
import numpy as np

"""
This program reads dataset file 'data.xlsx' in current directory and converts it into
pandas dataframe. If the file is not found or its name is wrong when reading it,
then raise Exception 'File file_name not found in path path/to/file'.

First this program checks if there exists directory named 'results', if not then
it creates one and then starts data handling.

The main funtion calls individual functions that each one creates plots 
to the 'results' folder as a png. format.


"""

def check_results_dir():
    # Check if results directory exists, if not then create it. 
    # Otherwise do nothing.
    current_dir = os.getcwd()
    results_dir = os.path.join(current_dir, "results")

    if not os.path.isdir(results_dir):
        os.mkdir(results_dir)


def read_dataframe(file_name):
    # Read dataframe from given excel file format .xlsx
    # If file is found then return dataframe.
    # If file is not found, raise exception that file is not found.
    current_dir = os.getcwd()
    file_path = os.path.join(current_dir, file_name)
    if os.path.exists(file_path):
        df = pd.read_excel(file_name, engine = "openpyxl")
        return df

    else:
        raise Exception(f"File {file_name} not found in path {file_path}")


def iterations_and_velocity(won):

    # read data from excel to datadrame
    df = read_dataframe("data.xlsx")

    # decide if want to look only "Won" status or all companies 
    # won = True/False
    if won:
        df_won = df[df["Opportunity Status"] == "Won"]

        velocity_won = df_won["Sales Velocity"].to_list()
        iterations_won = df_won["Sales Stage Iterations"].to_list()

        # plot scatter plot where x-axis is iterations and y-axis is sales velocity
        plt.ylim(0, 150)
        plt.scatter(x = iterations_won, y = velocity_won)
        plt.xlabel("Iterations")
        plt.ylabel("Sales velocity")
        plt.title("Iteration vs sales velocity in all companies")
        
        # Save figure to results folder
        current_dir = os.getcwd()
        save_path = os.path.join(current_dir, "results\\iterations_vs_sales_won_status.png")
        plt.savefig(save_path)
        
        #plt.show()

    # read velocity and iterations columns to lists
    velocity = df["Sales Velocity"].to_list()
    iterations = df["Sales Stage Iterations"].to_list()

    # plot scatter plot where x-axis is iterations and y-axis is sales velocity
    plt.ylim(0, 150)
    plt.scatter(x = iterations, y = velocity)
    plt.xlabel("Iterations")
    plt.ylabel("Sales velocity")
    plt.title("Iteration vs sales velocity in all companies")
    
    # Save figure to results folder
    current_dir = os.getcwd()
    save_path = os.path.join(current_dir, "results\\iterations_vs_sales.png")
    plt.savefig(save_path)
    
    #plt.show()


def tech_vs_erp():
    # read data from excel to datadrame
    df = read_dataframe("data.xlsx")

    # Separate tech and ERP into different dataframes
    tech_df = df[df["Technology\nPrimary"] == "Technical Business Solutions"]
    erp_df = df[df["Technology\nPrimary"] == "ERP Implementation"]

    # tech velocity and iterations
    tech_velocity = tech_df["Sales Velocity"].to_list()
    tech_iterations = tech_df["Sales Stage Iterations"].to_list()

    # ERP velocity and iterations
    erp_velocity = erp_df["Sales Velocity"].to_list()
    erp_iterations = erp_df["Sales Stage Iterations"].to_list()


    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2)
    #fig.suptitle('Tech vs ERP in iteration vs sales velocity')
    ax1.hist(tech_velocity, bins = 30)
    ax2.hist(erp_velocity, bins = 30)

    ax1.set_xlim(0, 150)
    ax2.set_xlim(0, 150)

    ax1.set_ylim(0, 6000)
    ax2.set_ylim(0, 6000)

    ax1.set_title("TECH sales velocity")
    ax2.set_title("ERP sales velocity")

    ax3.set_ylim(0, 150)
    ax3.scatter(x = tech_iterations, y = tech_velocity)
    ax3.set_xlabel("Tech Iterations")
    ax3.set_ylabel("Tech Sales velocity")

    ax4.set_ylim(0, 150)
    ax4.scatter(x = erp_iterations, y = erp_velocity)
    ax4.set_xlabel("ERP Iterations")
    ax4.set_ylabel("ERP Sales velocity")

    plt.tight_layout()

    # Save figure to results folder
    current_dir = os.getcwd()
    save_path = os.path.join(current_dir, "results\\Tech_vs_ERP.png")
    plt.savefig(save_path)

    #plt.show()

def erp_enterprise_vs_marketing_sales():
    # read data from excel to datadrame
    df = read_dataframe("data.xlsx")
    
    erp_df = df[df["Technology\nPrimary"] == "ERP Implementation"]

    marketing_df = erp_df[erp_df["B2B Sales Medium"] == "Marketing"]

    enterprise_df = erp_df[erp_df["B2B Sales Medium"] == "Enterprise Sellers"]

    marketing_vel = marketing_df["Sales Velocity"].to_list()
    enterprise_vel = enterprise_df["Sales Velocity"].to_list()

    fig, (ax1, ax2) = plt.subplots(1, 2)

    ax1.hist(marketing_vel, bins = 30)
    ax2.hist(enterprise_vel, bins = 30)

    ax1.set_xlim(0, 150)
    ax2.set_xlim(0, 150)

    ax1.set_ylim(0, 3000)
    ax2.set_ylim(0, 3000)

    ax1.set_title("Marketing sales velocity")
    ax2.set_title("Enterprise Sellers sales velocity")

    fig.supxlabel("Sales velocity")

    plt.tight_layout()

    # Save figure to results folder
    current_dir = os.getcwd()
    save_path = os.path.join(current_dir, "results\\ERP_enterprise_vs_marketing_sales_velocity.png")
    plt.savefig(save_path)

    # plt.show()


def erp_enterprise_vs_marketing_iterations():
    # read data from excel to datadrame
    df = read_dataframe("data.xlsx")
    
    erp_df = df[df["Technology\nPrimary"] == "ERP Implementation"]

    marketing_df = erp_df[erp_df["B2B Sales Medium"] == "Marketing"]

    enterprise_df = erp_df[erp_df["B2B Sales Medium"] == "Enterprise Sellers"]

    marketing_iter = marketing_df["Sales Stage Iterations"].to_list()
    enterprise_iter = enterprise_df["Sales Stage Iterations"].to_list()

    fig, (ax1, ax2) = plt.subplots(1, 2)

    ax1.hist(marketing_iter, bins = 15)
    ax2.hist(enterprise_iter, bins = 15)

    ax1.set_xlim(0, 17)
    ax2.set_xlim(0, 17)

    ax1.set_ylim(0, 12000)
    ax2.set_ylim(0, 12000)

    ax1.set_title("ERP Marketing sales iterations")
    ax2.set_title("ERP Enterprise Sellers sales iterations")
    
    fig.supxlabel("Iterations")

    plt.tight_layout()

    # Save figure to results folder
    current_dir = os.getcwd()
    save_path = os.path.join(current_dir, "results\\ERP_enterprise_vs_marketing_sales_iterations.png")
    plt.savefig(save_path)



def main():
    check_results_dir()
    iterations_and_velocity(won = False)
    tech_vs_erp()
    erp_enterprise_vs_marketing_sales()
    erp_enterprise_vs_marketing_iterations()


if __name__ == "__main__":
    main()
