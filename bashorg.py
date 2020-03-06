import requests
from lxml import html

def get_bashorg():
    r = requests.get('http://bash.im/random')
    tree = html.fromstring(r.content)
    cittext = tree.xpath('/html/body/div[1]/main/section/article[1]/div/div/text()')
    rettext = ''
    for i in cittext:
        rettext += i + '\n'
    return rettext

if __name__ == "__main__":
    text = get_bashorg()
    print(text)