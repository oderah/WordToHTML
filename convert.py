import mammoth
from bs4 import BeautifulSoup
import re
import math

style_map = """
u =>
table => div.table-responsive > table.table.table-condensed.table-bordered.table-responsive
"""

toc_prefix = '<p><a href="#_Toc'
end_of_toc = '</p>'

def hc_sc_clean(result, lang='en'):
    # table of content
    index = result.find(toc_prefix)
    result = result[:index] + '<ul>' + result[index:]
    index = result.rfind(toc_prefix)
    index = result.find(end_of_toc, index)
    index += 4
    result = result[:index] + '</ul>' + result[index:]

    # substitutions
    result = re.sub(
        '<p>\s{0,}<a href="#_Toc.*?">(.*?)</a>\s{0,}</p>',
        r'<li><a href="#a">\1</a></li>',
        result
    )
    result =  re.sub('<thead><tr>', '<thead><tr class="bg-primary">', result)
    result = re.sub('<h2>', '<h2 id="a">', result)
    result = re.sub('<h3>', '<h3 id="a">', result)
    result = re.sub('<a id=".*?">\s{0,}</a>', '', result)
    result = re.sub('<th><p>(.*?)</p></th>', r'<th>\1</th>', result)
    result = re.sub('<td><p>(.*?)</p></td>', r'<td>\1</td>', result)
    result = re.sub(
        '<sup>{0,1}<a href="#endnote-(\d+)" id="endnote-ref-\d+">.*?</a></sup>{0,1}',
        r'<sup><a class="fn-lnk" href="#fn\1"><span class="wb-inv">Footnote </span>\1</a></sup>' \
            if lang == 'en' \
            else r'<sup><a class="fn-lnk" href="#fn\1"><span class="wb-inv">Note de bas de page </span>\1</a></sup>',
        result
    )

    # cleanup
    result = BeautifulSoup(result, features="html.parser") \
        .prettify()

    return result

def export(result, path):
    # chunk
    n = math.ceil(len(result) / 4)
    chunks = [ result[i:i+n] for i in range(0, len(result), n) ]

    # export
    for index, chunk in enumerate(chunks):
        chunk_path = path + f'-part{ index + 1 }.html'
        with open(chunk_path, "w") as html_file:
            html_file.write(chunk)

    print(f"Successfully converted { len(chunks) } parts to => 'converted/'")


def main():
    lang = input("Language (en|fr) =>")
    file_to_convert = input("File to convert (include extension) =>")
    file_path = f"./doc_files/{ file_to_convert }"
    converted_file = file_to_convert.split(".doc")[0]
    converted_path = f"./converted/{ converted_file }"

    # format
    with open(file_path, "rb") as docx_file:
        result = mammoth.convert_to_html(
            docx_file,
            style_map=style_map
        ).value

    result = hc_sc_clean(result, lang)

    # export
    export(result, converted_path)

if __name__ == '__main__':
    main()
