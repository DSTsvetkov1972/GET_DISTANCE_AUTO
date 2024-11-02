import markdown, os

with open('README.md', 'r', encoding='utf-8') as f:
    markdown_string = f.read()

html_string = markdown.markdown(markdown_string)

with open('READMY.html', 'w') as f:
    f.write(html_string)

os.startfile('sample.html')