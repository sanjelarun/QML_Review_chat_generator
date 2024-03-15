import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def save_plot_publications_by_year(data):
    # Clean non-integer values and convert to integer
    data['Publication Year'] = pd.to_numeric(data['Publication Year'], errors='coerce').dropna().astype(int)
    
    plt.figure(figsize=(13, 9))
    year_counts = data['Publication Year'].value_counts().sort_values(ascending=False)
    ax = sns.barplot(x=year_counts.index, y=year_counts.values, color='blue', order=year_counts.index)
    
    plt.ylabel("Number of Papers", fontsize=16)
    plt.xlabel("Published Year", fontsize=16)
    plt.xticks(rotation=45, fontsize=16)
    ax.set_xticklabels(year_counts.index.astype(int))
    ax.yaxis.grid(True, linestyle='--', linewidth=0.5, alpha=0.7)
    plt.tight_layout()
    plt.savefig("publications_by_year.png", dpi=500)
    plt.close()

def save_plot_papers_by_type(data):
    plt.figure(figsize=(12, 8))
    ax = sns.countplot(data=data, x='Item Type', order=data['Item Type'].value_counts().index, color="blue")
    plt.ylabel("Number of Papers", fontsize=16)
    plt.xlabel("Publication Type", fontsize=16)
    plt.xticks(rotation=45, fontsize=16)
    
    # Add horizontal grid lines
    ax.yaxis.grid(True, linestyle='--', linewidth=0.5, alpha=0.7)
    
    plt.tight_layout()
    plt.savefig("papers_by_type.png", dpi=500)
    plt.close()

def save_plot_authors_per_paper(data):
    data['Author Count'] = data['Author'].apply(lambda x: len(str(x).split(";")) if pd.notnull(x) else 0)
    author_count_counts = data['Author Count'].value_counts().sort_values(ascending=False)
    plt.figure(figsize=(12, 8))
    ax = sns.barplot(x=author_count_counts.index, y=author_count_counts.values, color="blue", order=author_count_counts.index)
    plt.ylabel("Number of Papers", fontsize=16)
    plt.xlabel("Number of Authors", fontsize=16)
    
    # Add horizontal grid lines
    ax.yaxis.grid(True, linestyle='--', linewidth=0.5, alpha=0.7)
    
    plt.tight_layout()
    plt.savefig("authors_per_paper.png", dpi=500)
    plt.close()

def save_plot_top_publishers(data):
    top_publishers = data['Publication Title'].value_counts().head(10).index
    filtered_data = data[data['Publication Title'].isin(top_publishers)]
    plt.figure(figsize=(12, 8))
    sns.countplot(data=filtered_data, y='Publication Title', order=top_publishers, color="blue")
    plt.xlabel("Number of Papers", fontsize=16)
    plt.ylabel("Publisher", fontsize=16)
    plt.tight_layout()
    plt.savefig("top_publishers.png", dpi=500)
    plt.close()

def save_driver_function():
    data = pd.read_excel("mapping_study_latest-bikram.xlsx", sheet_name="AllPapers")
    data.columns = data.iloc[0]
    data = data.drop(0).reset_index(drop=True).dropna(how="all")
    save_plot_publications_by_year(data)
    save_plot_papers_by_type(data)
    save_plot_authors_per_paper(data)
    save_plot_top_publishers(data)

# Call the save_driver_function to generate and save all the plots
save_driver_function()
