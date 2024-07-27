import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog, messagebox, StringVar
from tkinter import ttk
from matplotlib.backends.backend_pdf import PdfPages
import seaborn as sns
from wordcloud import WordCloud
from collections import Counter
from itertools import combinations
import networkx as nx
from matplotlib.backends.backend_pdf import PdfPages
from PyPDF2 import PdfWriter, PdfReader
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Set the default font to Times New Roman for matplotlib
plt.rcParams['font.family'] = 'Times New Roman'
plt.rcParams['font.sans-serif'] = ['Times New Roman']

class BibliometricsApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Bibliometrics Analysis Tool")

        # Create widgets
        self.label = ttk.Label(master, text="Select a CSV file with Scopus data:")
        self.label.grid(row=0, column=0, padx=10, pady=10)

        self.load_button = ttk.Button(master, text="Load Data", command=self.load_data)
        self.load_button.grid(row=1, column=0, padx=10, pady=10)

        self.year_label = ttk.Label(master, text="Select Start Year:")
        self.year_label.grid(row=2, column=0, padx=10, pady=10)

        self.year_var = StringVar()
        self.year_combobox = ttk.Combobox(master, textvariable=self.year_var, state="readonly")
        self.year_combobox.grid(row=3, column=0, padx=10, pady=10)

        self.analyze_button = ttk.Button(master, text="Analyze Data", command=self.analyze_data)
        self.analyze_button.grid(row=4, column=0, padx=10, pady=10)
        self.analyze_button["state"] = "disabled"

        self.save_button = ttk.Button(master, text="Save Analysis", command=self.save_analysis)
        self.save_button.grid(row=5, column=0, padx=10, pady=10)
        self.save_button["state"] = "disabled"

        self.df = None
        self.figures = []
        self.tables = []

    def load_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return

        try:
            self.df = pd.read_csv(file_path)

            # Check if required columns exist
            required_columns = ['Authors', 'Author full names', 'Author(s) ID', 'Title', 'Year', 'Source title', 'Volume', 'Issue', 'Page start', 'Page end', 'Cited by', 'Affiliations', 'Author Keywords']
            if not all(col in self.df.columns for col in required_columns):
                raise ValueError(f"The CSV file must contain the following columns: {', '.join(required_columns)}")

            # Populate year combobox
            years = self.df['Year'].dropna().unique()
            self.year_combobox['values'] = sorted(years)
            self.year_combobox.current(0)

            messagebox.showinfo("Success", "Data loaded successfully!")
            self.analyze_button["state"] = "normal"

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def analyze_data(self):
        if self.df is None:
            messagebox.showerror("Error", "No data loaded. Please load data first.")
            return

        try:
            start_year = int(self.year_var.get())
            filtered_df = self.df[self.df['Year'] >= start_year]

            if filtered_df.empty:
                messagebox.showerror("Error", f"No data available from the year {start_year} onwards.")
                return

            # Clear previous figures and tables
            self.figures.clear()
            self.tables.clear()

            # Basic Information
            min_year = filtered_df['Year'].min()
            max_year = filtered_df['Year'].max()
            unique_authors = filtered_df['Authors'].dropna().str.split(',', expand=True).stack().str.strip().unique()
            total_authors = len(unique_authors)

            # Analysis
            publication_count_by_year = filtered_df['Year'].value_counts().sort_index()
            growth_percentage_by_year = publication_count_by_year.pct_change().fillna(0) * 100
            growth_percentage_by_year = growth_percentage_by_year.round(1)
            publications_per_source = filtered_df['Source title'].value_counts()
            cited_by_distribution = filtered_df['Cited by'].fillna(0).astype(float)
            average_citations_per_year = filtered_df.groupby('Year')['Cited by'].mean().fillna(0)

            # Correlation between Number of Citations and Number of Authors per Paper
            filtered_df['Number of Authors'] = filtered_df['Authors'].str.split(',').str.len()
            correlation = filtered_df[['Number of Authors', 'Cited by']].dropna().corr().iloc[0, 1]

            # Plot data
            sns.set(style="whitegrid")

            # Publication Count by Year
            fig1, ax1 = plt.subplots(figsize=(8.27, 11.69))  # A4 size
            sns.barplot(x=publication_count_by_year.index, y=publication_count_by_year.values, ax=ax1, palette='viridis')
            ax1.set_title('Publication Count by Year', fontsize=14, family='Times New Roman')
            ax1.set_xlabel('Year', fontsize=12, family='Times New Roman')
            ax1.set_ylabel('Number of Publications', fontsize=12, family='Times New Roman')
            self.figures.append(fig1)
            # Table for publication count by year with growth percentage
            fig1_table, ax1_table = plt.subplots(figsize=(8.27, 11.69))
            ax1_table.axis('off')
            table_data = pd.DataFrame({
                'Year': publication_count_by_year.index,
                'Number of Publications': publication_count_by_year.values,
                'Growth (%)': growth_percentage_by_year.values
            }).to_numpy()
            table = ax1_table.table(cellText=table_data, colLabels=['Year', 'Number of Publications', 'Growth (%)'], cellLoc='center', loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1.2, 1.2)  # Scale the table for better readability
            self.tables.append(fig1_table)

            # Number of Publications per Source Title
            fig2, ax2 = plt.subplots(figsize=(8.27, 11.69))
            sns.barplot(y=publications_per_source.index[:10], x=publications_per_source.values[:10], ax=ax2, palette='magma')
            ax2.set_title('Top 10 Sources by Number of Publications', fontsize=14, family='Times New Roman')
            ax2.set_xlabel('Number of Publications', fontsize=12, family='Times New Roman')
            ax2.set_ylabel('Source Title', fontsize=12, family='Times New Roman')
            self.figures.append(fig2)
            # Table for top sources
            fig2_table, ax2_table = plt.subplots(figsize=(8.27, 11.69))
            ax2_table.axis('off')
            table_data = publications_per_source.head(10).reset_index().rename(columns={'index': 'Source Title', 'Source title': 'Number of Publications'}).to_numpy()
            table = ax2_table.table(cellText=table_data, colLabels=['Source Title', 'Number of Publications'], cellLoc='center', loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1.2, 1.2)  # Scale the table for better readability
            self.tables.append(fig2_table)

            # Top Authors by Number of Publications
            fig3, ax3 = plt.subplots(figsize=(8.27, 11.69))
            top_authors = filtered_df['Authors'].dropna().str.split(',', expand=True).stack().str.strip().value_counts()
            sns.barplot(y=top_authors.index[:10], x=top_authors.values[:10], ax=ax3, palette='coolwarm')
            ax3.set_title('Top 10 Authors by Number of Publications', fontsize=14, family='Times New Roman')
            ax3.set_xlabel('Number of Publications', fontsize=12, family='Times New Roman')
            ax3.set_ylabel('Author', fontsize=12, family='Times New Roman')
            self.figures.append(fig3)
            # Table for top authors
            fig3_table, ax3_table = plt.subplots(figsize=(8.27, 11.69))
            ax3_table.axis('off')
            table_data = top_authors.head(10).reset_index().rename(columns={'index': 'Author', 'Authors': 'Number of Publications'}).to_numpy()
            table = ax3_table.table(cellText=table_data, colLabels=['Author', 'Number of Publications'], cellLoc='center', loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1.2, 1.2)  # Scale the table for better readability
            self.tables.append(fig3_table)

            # Top Keywords by Coupling
            fig4, ax4 = plt.subplots(figsize=(8.27, 11.69))
            keywords = filtered_df['Author Keywords'].dropna().str.split(';').explode().str.strip().value_counts()
            sns.barplot(y=keywords.index[:10], x=keywords.values[:10], ax=ax4, palette='inferno')
            ax4.set_title('Top 10 Keywords by Coupling', fontsize=14, family='Times New Roman')
            ax4.set_xlabel('Number of Couplings', fontsize=12, family='Times New Roman')
            ax4.set_ylabel('Keyword', fontsize=12, family='Times New Roman')
            self.figures.append(fig4)
            # Table for top keywords
            fig4_table, ax4_table = plt.subplots(figsize=(8.27, 11.69))
            ax4_table.axis('off')
            table_data = keywords.head(10).reset_index().rename(columns={'index': 'Keyword', 'Author Keywords': 'Number of Couplings'}).to_numpy()
            table = ax4_table.table(cellText=table_data, colLabels=['Keyword', 'Number of Couplings'], cellLoc='center', loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1.2, 1.2)  # Scale the table for better readability
            self.tables.append(fig4_table)

            # Conceptual Structure (Word Cloud)
            fig5, ax5 = plt.subplots(figsize=(8.27, 11.69))
            text = ' '.join(filtered_df['Title'].dropna())
            wordcloud = WordCloud(width=800, height=1200, background_color='white').generate(text)
            ax5.imshow(wordcloud, interpolation='bilinear')
            ax5.axis('off')
            ax5.set_title('Conceptual Structure (Word Cloud)', fontsize=14, family='Times New Roman')
            self.figures.append(fig5)

            # Top Authors by Citations
            fig6, ax6 = plt.subplots(figsize=(8.27, 11.69))
            top_authors_cited = filtered_df.groupby('Authors')['Cited by'].sum().sort_values(ascending=False)
            sns.barplot(y=top_authors_cited.index[:10], x=top_authors_cited.values[:10], ax=ax6, palette='cool')
            ax6.set_title('Top 10 Authors by Citations', fontsize=14, family='Times New Roman')
            ax6.set_xlabel('Number of Citations', fontsize=12, family='Times New Roman')
            ax6.set_ylabel('Author', fontsize=12, family='Times New Roman')
            self.figures.append(fig6)
            # Table for top authors by citations
            fig6_table, ax6_table = plt.subplots(figsize=(8.27, 11.69))
            ax6_table.axis('off')
            table_data = top_authors_cited.head(10).reset_index().rename(columns={'index': 'Author', 'Cited by': 'Number of Citations'}).to_numpy()
            table = ax6_table.table(cellText=table_data, colLabels=['Author', 'Number of Citations'], cellLoc='center', loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1.2, 1.2)  # Scale the table for better readability
            self.tables.append(fig6_table)

            # Social Structure (Co-authorship Network)
            fig7, ax7 = plt.subplots(figsize=(8.27, 11.69))
            coauthor_pairs = filtered_df['Authors'].dropna().str.split(',').apply(lambda x: list(combinations([author.strip() for author in x], 2)))
            coauthor_pairs_flat = [pair for sublist in coauthor_pairs for pair in sublist]
            coauthor_network = Counter(coauthor_pairs_flat)
            G = nx.Graph()
            for (author1, author2), weight in coauthor_network.items():
                G.add_edge(author1, author2, weight=weight)
            pos = nx.spring_layout(G, k=0.1)
            nx.draw(G, pos, with_labels=True, node_size=50, font_size=10, ax=ax7)
            ax7.set_title('Social Structure (Co-authorship Network)', fontsize=14, family='Times New Roman')
            self.figures.append(fig7)

            # Create a summary table for the first page
            summary_table_data = [
                ['Time Period', f"{min_year} - {max_year}"],
                ['Number of Unique Authors', total_authors],
                ['Correlation between Number of Citations and Number of Authors', f"{correlation:.2f}"]
            ]
            fig_summary, ax_summary = plt.subplots(figsize=(8.27, 11.69))
            ax_summary.axis('off')
            summary_table = ax_summary.table(cellText=summary_table_data, colLabels=['Metric', 'Value'], cellLoc='center', loc='center')
            summary_table.auto_set_font_size(False)
            summary_table.set_fontsize(12)
            summary_table.scale(1.2, 1.2)  # Scale the table for better readability
            self.tables.insert(0, fig_summary)

            self.save_button["state"] = "normal"

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def save_analysis(self):
        if not self.figures and not self.tables:
            messagebox.showerror("Error", "No analysis to save. Please perform analysis first.")
            return

        try:
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if not save_path:
                return

            with PdfPages(save_path) as pdf:
                # Add index page
                index_page = plt.figure(figsize=(8.27, 11.69))  # A4 size
                ax_index = index_page.add_subplot(111)
                ax_index.axis('off')
                index_text = [
                    "Index",
                    "1. Summary Information - Page 1",
                    "2. Publication Count by Year - Page 2",
                    "3. Top 10 Sources by Number of Publications - Page 3",
                    "4. Top 10 Authors by Number of Publications - Page 4",
                    "5. Top 10 Keywords by Coupling - Page 5",
                    "6. Conceptual Structure (Word Cloud) - Page 6",
                    "7. Top 10 Authors by Citations - Page 7",
                    "8. Social Structure (Co-authorship Network) - Page 8"
                ]
                for line in index_text:
                    ax_index.text(0.1, 1 - 0.1 * index_text.index(line), line, fontsize=12, family='Times New Roman')
                pdf.savefig(index_page, bbox_inches='tight')

                # Add summary page
                pdf.savefig(self.tables[0], bbox_inches='tight')

                # Add all figures and tables
                for fig in self.figures:
                    pdf.savefig(fig, bbox_inches='tight')
                for table in self.tables[1:]:
                    pdf.savefig(table, bbox_inches='tight')

            # Add page numbers and footer note using PyPDF2
            writer = PdfWriter()
            with open(save_path, "rb") as f:
                reader = PdfReader(f)
                for i in range(len(reader.pages)):
                    page = reader.pages[i]
                    # Create a new page to overlay the footer and page number
                    packet = BytesIO()
                    c = canvas.Canvas(packet, pagesize=letter)
                    c.setFont("Times-Roman", 8)
                    c.drawString(200, 10, f"Data Analysis done through Bibliometrics App by Shivam Moradia")
                    c.drawString(580, 10, f"Page {i + 1}")
                    c.save()
                    
                    packet.seek(0)
                    new_pdf = PdfReader(packet)
                    overlay = new_pdf.pages[0]
                    page.merge_page(overlay)
                    writer.add_page(page)
            
            with open(save_path, "wb") as f:
                writer.write(f)

            messagebox.showinfo("Success", "Analysis saved successfully!")

        except Exception as e:
            messagebox.showerror("Error", str(e))

# Create the main window
root = Tk()
app = BibliometricsApp(root)
root.mainloop()