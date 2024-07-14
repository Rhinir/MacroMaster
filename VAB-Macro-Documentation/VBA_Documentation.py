import os
import cohere
import markdown2
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib import colors
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains


COHERE_API_KEY_TEXT = "cfsUYJQURZuHCThMSSUEUWWv3LVdTpcdP7YXH8RS"
os.environ["COHERE_API_KEY"] = COHERE_API_KEY_TEXT


def auto_documentation(vba_macro1):
    
    prompt = f"""You are a tool built for taking VBA Macros as input and creating documentation of the underlying logic, the data flow and the process flow of the macro code. Also, create a suitable title for the documentation.
    The VBA Macro is: {vba_macro1}
    The required fields are: Title, Underlying Logic, Data Flow, Process Flow
    Extract the required fields from the provided lines of code.
    Give the output in the form of a markdown file.
    Don't add any other lines in the response except the required fields"""
    
    COHERE_API_KEY = "cfsUYJQURZuHCThMSSUEUWWv3LVdTpcdP7YXH8RS"
    co = cohere.Client(COHERE_API_KEY)
    response = co.chat(model="command-r", message=prompt, max_tokens=1000)


    with open("Documentation.md", "w") as file:
        file.write(response.text + "\n")
    
    with open("Documentation.md", 'r', encoding='utf-8') as f:
        markdown_content = f.read()

    html_content = markdown2.markdown(markdown_content)
    # print(html_content)

    output_pdf = "Documentation.pdf"
    doc = SimpleDocTemplate(output_pdf, pagesize=letter)
    story = []

    paragraphs = html_content.split('<h')
    for paragraph in paragraphs:
        if paragraph.startswith('1>'):
            title_style = ParagraphStyle(
                name='Title',
                fontSize=22,
                leading=24,
                alignment=TA_CENTER,
            )
            title_paragraph = Paragraph(paragraph[2:], title_style)
            story.append(title_paragraph)
            # story.append(vba_macro1)
            story.append(Spacer(1, 15))
        elif paragraph.startswith('2>'):
            heading_style = ParagraphStyle(
                name='Heading',
                fontSize=16,
                leading=20,
                alignment=TA_LEFT,
                textColor=colors.blue,
            )
            heading_paragraph = Paragraph(paragraph[2: paragraph.index('</h2>')], heading_style)
            story.append(heading_paragraph)
            body_style = ParagraphStyle(
                name='BulletPoints',
                fontSize=12,
                leading=16,
                bulletIndent=20,
                leftIndent=20,
            )
            for point in paragraph[paragraph.index('</h2>')+5:].split('.'):
                body_paragraph = Paragraph(point, body_style, bulletText='•')
                story.append(body_paragraph)
                story.append(Spacer(1, 6))
        else:
            body_style = ParagraphStyle(
                name='BulletPoints',
                fontSize=12,
                leading=16,
                bulletIndent=20,
                leftIndent=20,
            )
            for point in paragraph.split('.'):
                body_paragraph = Paragraph(point, body_style, bulletText='•')
                story.append(body_paragraph)
                story.append(Spacer(1, 6))

    doc.build(story)

def functional_logic_extractor(vba_macro1):
    
    prompt = f"""You are a tool built for taking VBA Macros as input and extract and explain the functional logic embedded within the VBA macro. The explanation should be in a manner any non-technical stakeholder can easily understand. Also, create a suitable title for the code. 
    The VBA Macro is: {vba_macro1}
    The required fields are: Title, Functional Logic
    Extract the required field from the provided lines of code.
    Give the output in the form of a markdown file
    Don't add any other lines in the response except the required fields"""
    
    COHERE_API_KEY = "cfsUYJQURZuHCThMSSUEUWWv3LVdTpcdP7YXH8RS"
    co = cohere.Client(COHERE_API_KEY)
    response = co.chat(model="command-r", message=prompt, max_tokens=1000)


    with open("Functional_Logic.md", "w") as file:
        file.write(response.text + "\n")

    with open("Functional_Logic.md", 'r', encoding='utf-8') as f:
        markdown_content = f.read()

    html_content = markdown2.markdown(markdown_content)
    # print(html_content)

    output_pdf = "Functional_Logic.pdf"
    doc = SimpleDocTemplate(output_pdf, pagesize=letter)
    story = []

    paragraphs = html_content.split('<h')
    for paragraph in paragraphs:
        if paragraph.startswith('1>'):
            title_style = ParagraphStyle(
                name='Title',
                fontSize=22,
                leading=24,
                alignment=TA_CENTER,
            )
            title_paragraph = Paragraph(paragraph[2:], title_style)
            story.append(title_paragraph)
            story.append(Spacer(1, 15))
        elif paragraph.startswith('2>'):
            heading_style = ParagraphStyle(
                name='Heading',
                fontSize=16,
                leading=20,
                alignment=TA_LEFT,
                textColor=colors.blue,
            )
            heading_paragraph = Paragraph(paragraph[2: paragraph.index('</h2>')], heading_style)
            story.append(heading_paragraph)
            body_style = ParagraphStyle(
                name='BulletPoints',
                fontSize=12,
                leading=16,
                bulletIndent=20,
                leftIndent=20,
            )
            for point in paragraph[paragraph.index('</h2>')+5:].split('.'):
                body_paragraph = Paragraph(point, body_style, bulletText='•')
                story.append(body_paragraph)
                story.append(Spacer(1, 6))
        else:
            body_style = ParagraphStyle(
                name='BulletPoints',
                fontSize=12,
                leading=16,
                bulletIndent=20,
                leftIndent=20,
            )
            for point in paragraph.split('.'):
                body_paragraph = Paragraph(point, body_style, bulletText='•')
                story.append(body_paragraph)
                story.append(Spacer(1, 6))
    doc.build(story)

def create_flowchart(vba_macro1):
    prompt = f"""You are a tool built for taking a VBA Macro code as input and write an equivalent python file which does the same job and has the same process flow. 
    The VBA Macro is: {vba_macro1}
    The required fields are: Python File
    Extract the required field from the provided lines of code.
    Give the output in the form of a python file.
    Don't add any other lines in the response except the required fields
    Dont add ''' and python at the beginning"""
    
    COHERE_API_KEY = "cfsUYJQURZuHCThMSSUEUWWv3LVdTpcdP7YXH8RS"
    co = cohere.Client(COHERE_API_KEY)
    response = co.chat(model="command-r", message=prompt, max_tokens=1000)


    with open("Python_Code.py", "w") as file:
        file.write(response.text + "\n")

    os.system("python -m pyflowchart Python_Code.py -o flowchart.html")

    # options = webdriver.EdgeOptions()

    # cwd = os.getcwd()
    # prefs = {"download.default_directory" : cwd}
    # options.add_experimental_option("prefs", prefs)
    # options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # options.use_chromium = True
    # options.add_argument("headless")
    
    # browser = webdriver.Edge(options = options)
    # browser.get("C://Users//HP//Desktop//SG-Hackathon//SG-Hackathon//VAB-Macro-Documentation//flowchart.html")

    # link = browser.find_element(By.XPATH, '/html/body/div[5]/a')
    # actions = ActionChains(browser)
    # actions.move_to_element(link).click().perform()

    # time.sleep(10)
    # browser.close()

def modernize_code(vba_macro1):
    prompt = f"""You are a tool built for taking a VBA Macro code as input and write an python file which does the same job but is more efficient and optimized. 
    The VBA Macro is: {vba_macro1}
    The required fields are: Python File
    Extract the required field from the provided lines of code.
    Give the output in the form of a python file.
    Don't add any other lines in the response except the required fields
    Dont add ''' and python at the beginning"""
    
    COHERE_API_KEY = "cfsUYJQURZuHCThMSSUEUWWv3LVdTpcdP7YXH8RS"
    co = cohere.Client(COHERE_API_KEY)
    response = co.chat(model="command-r", message=prompt, max_tokens=1000)


    with open("Modernized_Code.py", "w") as file:
        file.write(response.text + "\n")

def optimize_code(vba_macro1):
    prompt = f"""You are a tool built for taking a VBA Macro code as input and write an optimized version of VBA Macro Code. Only if the input code is wrong somewhere or not taking an optimized way, specify that along with that part of the code, which means specify the limitations of the input code if any. 
    The VBA Macro is: {vba_macro1}
    The required fields are: Limitations, Optimized Code
    Extract the required fields from the provided lines of code.
    Give the output in the form of a text file.
    Don't add any other lines in the response except the required fields"""
     
    COHERE_API_KEY = "cfsUYJQURZuHCThMSSUEUWWv3LVdTpcdP7YXH8RS"
    co = cohere.Client(COHERE_API_KEY)
    response = co.chat(model="command-r", message=prompt, max_tokens=1000)


    with open("Optimized_Code.txt", "w") as file:
        file.write(response.text + "\n")
