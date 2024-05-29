from reportlab.lib.pagesizes import letter 
from reportlab.lib import colors 
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd 
import yfinance as yf

ticker = 'RMS.PA'

company = yf.Ticker(ticker)
company_data = company.info

pdf_filename = f"{ticker}_report.pdf"
doc = SimpleDocTemplate(pdf_filename, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
styles = getSampleStyleSheet()
elements = []

styles.add(ParagraphStyle(name='CustomTitle', fontSize=20, textColor=colors.black, spaceAfter=30, alignment=1, fontName='Helvetica-Bold'))
styles.add(ParagraphStyle(name='CustomHeading2', fontSize=16, textColor=colors.black, spaceAfter=16, fontName='Helvetica-Bold'))

if 'BodyText' not in styles:
    styles.add(ParagraphStyle(name='BodyText', fontSize=12, textColor=colors.black, spaceAfter=10))
    
def create_styled_table(data):
    table = Table(data, colWidths=[doc.width / 2, doc.width / 2])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D3D3D3')), 
    ]))
    return table

def add_bordered_section(title, content_elements, elements, styles, bordered=True):
    title_paragraph = Paragraph(title, styles['CustomHeading2'])
    if bordered:
        bordered_table = Table(
            [[title_paragraph]] + [[content] for content in content_elements], 
            colWidths = [doc.width]
        )
        bordered_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BOX', (0, 0), (-1, -1), 2, colors.black),
            ('TOPPADDING', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FFD1B9')),
        ]))
        elements.append(bordered_table)
    else:
        elements.append(title_paragraph)
        elements.extend(content_elements)
    elements.append(Spacer(1, 14))
    
elements.append(Paragraph(f"Company Report : {company_data.get('longName', 'N/A')} ({ticker})", styles['CustomTitle']))
elements.append(Spacer(1, 30))

general_info = [
    ['Statement', 'Details'], 
    ['Full Name', company_data.get('longName', 'N/A')], 
    ['Sector', company_data.get('sector', 'N/A')], 
    ['Industry', company_data.get('industry', 'N/A')], 
    ['WebSite', company_data.get('website', 'N/A')], 
    ['Address', company_data.get('address1', 'N/A')], 
    ['City', company_data.get('city', 'N/A')], 
    ['Postal Code', company_data.get('zip', 'N/A')], 
    ['Country', company_data.get('country', 'N/A')], 
    ['Phone', company_data.get('phone', 'N/A')], 
    ['CEO', next((officer.get('name') for officer in company_data.get('companyOfficers', []) if officer.get('title') == 'Executive Chairman & CEO'), 'N/A')], 
    ['Number of Full-Time Employees', company_data.get('fullTimeEmployees', 'N/A')], 
]

general_info_elements = [create_styled_table(general_info)]
add_bordered_section("Company General Information:", general_info_elements, elements, styles, bordered=False)

description_elements = [Paragraph(company_data.get('longBusinessSummary', 'N/A'), styles['BodyText'])]
add_bordered_section("Company Description:", description_elements, elements, styles)

elements.append(PageBreak())

financial_data = [
    ['Statement', 'Details'], 
    ['Current Price', f"{company_data.get('currentPrice', 'N/A')} EUR"], 
    ['PER (forward)', company_data.get('forwardPE', 'N/A')], 
    ['PER (trailing)', company_data.get('trailingPE', 'N/A')], 
    ['Dividend Yield', f"{company_data.get('dividendYield', 0) * 100:.2f} %"], 
    ['Dividend Amount', f"{company_data.get('dividendRate', 'N/A')} EUR"], 
    ['Total Cash Per Share', f"{company_data.get('totalCashPerShare', 'N/A')} EUR"], 
]

financial_data_elements = [create_styled_table(financial_data)]
add_bordered_section("Key Financial Data:", financial_data_elements, elements, styles, bordered=False)

ratios_data = [
    ['Statement', 'Details'], 
    ['Beta', company_data.get('beta', 'N/A')], 
    ['Price-to-Book Ratio', company_data.get('priceToBook', 'N/A')], 
    ['PEG Ratio', company_data.get('pegRatio', 'N/A')], 
    ['Return on Assets (ROA)', f"{company_data.get('returnOnAssets', 0) * 100:.2f} %"], 
    ['Return on Equity (ROE)', f"{company_data.get('returnOnEquity', 0) * 100:.2f} %"], 
    ['Gross Margin', f"{company_data.get('grossMargins', 0) * 100:.2f} %"], 
    ['EBITDA Margin', f"{company_data.get('ebitdaMargins', 0) * 100:.2f} %"],  
    ['Operating Margin', f"{company_data.get('operatingMargins', 0) * 100:.2f} %"],  
]

ratios_data_elements = [create_styled_table(ratios_data)]
add_bordered_section("Financial Ratios:", ratios_data_elements, elements, styles, bordered=False)
    
analyst_data = [
    ['Statement', 'Details'], 
    ['Current Price', f"{company_data.get('currentPrice', 'N/A')} EUR"], 
    ['Target High Price', f"{company_data.get('targetHighPrice', 'N/A')} EUR"], 
    ['Target Low Price', f"{company_data.get('targetLowPrice', 'N/A')} EUR"], 
    ['Target Mean Price', f"{company_data.get('targetMeanPrice', 'N/A')} EUR"], 
    ['Target Median Price', f"{company_data.get('targetMedianPrice', 'N/A')} EUR"], 
    ['Recommendation Key', company_data.get('recommendationKey', 'N/A')], 
    ['Number of Analyst Opinions', company_data.get('numberOfAnalystOpinions', 'N/A')], 
]

boxplot_data = {
    'Current Price': company_data.get('currentPrice', 'N/A'), 
    'Target High Price': company_data.get('targetHighPrice', 'N/A'), 
    'Target Low Price': company_data.get('targetLowPrice', 'N/A'), 
    'Target Mean Price': company_data.get('targetMeanPrice', 'N/A'), 
    'Target Median Price': company_data.get('targetMedianPrice', 'N/A'), 
}

# Create boxplot
boxplot_df = pd.DataFrame(boxplot_data, index=[0])
plt.figure(figsize=(10, 6))
sns.boxplot(data=boxplot_df)
plt.title('Analyst Recommendations: Price Targets')
plt.ylabel('Price (EUR)')
plt.savefig('analyst_recommendations_boxplot.png')
plt.close()

boxplot_image_element = Image('analyst_recommendations_boxplot.png', width=doc.width * 0.8, height=300)
add_bordered_section("Analyst Recommendations: Price Targets (Box Plot)", [boxplot_image_element], elements, styles, bordered=False)

elements.append(PageBreak())

history = company.history(period="1y")
plt.figure(figsize=(12, 5))
plt.plot(history.index, history['Close'], label='Closing Price', color='blue')
plt.xlabel('Date')
plt.ylabel('Closing Price (EUR)')
plt.title('Price Trend (Past Year)')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig('price_trend.png')
plt.close()
price_trend_elements = [Image('price_trend.png', width=doc.width * 0.98, height=300)]
add_bordered_section("Price Trend (Past Year):", price_trend_elements, elements, styles)

plt.figure(figsize=(12, 5))
plt.bar(history.index, history['Volume'], label='Volume', color='green')
plt.xlabel('Date')
plt.ylabel('Volume')
plt.title('Volume Trend (Past Year)')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig('volume_trend.png')
plt.close()
volume_trend_elements = [Image('volume_trend.png', width=doc.width * 0.98, height=300)]
add_bordered_section("Volume Trend (Past Year):", volume_trend_elements, elements, styles)

def add_page_background(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(colors.HexColor("#FFDAB9"))
    canvas.rect(0, 0, doc.pagesize[0], doc.pagesize[1], fill=1)
    canvas.restoreState()


doc.build(elements, onFirstPage=add_page_background, onLaterPages=add_page_background)
print(f"PDF report '{pdf_filename}' has been created successfully.")






