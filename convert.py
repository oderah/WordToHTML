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
    print('Prepping TOC...')
    index = result.find(toc_prefix)
    result = result[:index] + '<ul>' + result[index:]
    index = result.rfind(toc_prefix)
    index = result.find(end_of_toc, index)
    index += 4
    result = result[:index] + '</ul>' + result[index:]

    # substitutions
    print('Making substitutions...')
    result = re.sub('<br />', '<br>', result)
    result = re.sub(
        'href="http://www.canada.ca/content/dam/|href="https://www.canada.ca/content/dam/ \
            |href="http://canada.ca/content/dam/|href="https://canada.ca/content/dam/ \
            |href="https://canada-preview.adobecqms.net/content/dam/',
        'href="/content/dam/',
        result
    )
    result = re.sub(
        'href="http://www.canada.ca/|href="https://www.canada.ca/|href="http://canada.ca/ \
            |href="https://canada.ca/|href="https://canada-preview.adobecqms.net/',
        'href="/content/canadasite/',
        result
    )
    result = re.sub('href="/(en|fr)/', r'href="/content/canadasite/\1/', result)
    result = re.sub('<p>\s{0,}(.*?)\s{0,}</p>', r'<p>\1</p>', result)
    result = re.sub(
        '<p>\s{0,}<a href="#_Toc.*?">(.*?)</a>\s{0,}</p>',
        r'<li><a href="#a">\1</a></li>',
        result
    )
    result = re.sub(
        '<table(.*?)><tr>',
        r'<table\1><caption class="text-left"></caption><thead><tr class="bg-primary">',
        result
    )
    result = re.sub('</table>', '</tbody></table>', result)
    result = re.sub('<thead><thead>', '<thead>', result)
    result = re.sub('</tbody></tbody>', '</tbody>', result)
    result = re.sub('<thead><tr>', '<thead><tr class="bg-primary">', result)
    result = re.sub('<h2>', '<h2 id="a">', result)
    result = re.sub('<h3>', '<h3 id="a">', result)
    result = re.sub('<a id=".*?">\s{0,}</a>', '', result)
    result = re.sub('<th(.*?)><p>(.*?)</p></th>', r'<th\1>\2</th>', result)
    result = re.sub('<td(.*?)><p>(.*?)</p></td>', r'<td\1>\2</td>', result)
    result = re.sub('href="file:.*?"', 'href="#"', result)
    result = re.sub('src="file:.*?"', 'src="#"', result)
    result = re.sub('<img.*?/>', '<img src="#" class="img-responsive" alt="" />', result)
    result = re.sub(
        '<sup>{0,1}<a href="#endnote-(\d+)" id="endnote-ref-\d+">.*?</a></sup>{0,1}',
        r'<sup><a class="fn-lnk" href="#fn\1"><span class="wb-inv">Footnote </span>\1</a></sup>' \
            if lang == 'en' \
            else r'<sup><a class="fn-lnk" href="#fn\1"><span class="wb-inv">Note de bas de page </span>\1</a></sup>',
        result
    )
    result = re.sub('<sup><sup>', '<sup>', result)
    result = re.sub('</sup></sup>', '</sup>', result)

    # cleanup
    # print('Prettifying...)
    # result = BeautifulSoup(result, 'html.parser') \
    #     .prettify()

    return result

def export(result, path, should_chunk=0):
    if should_chunk == 0:
        with open(f'{ path }.html', 'w') as html_file:
            html_file.write(result)
        print(f'Successfully converted file to => "{ path }.html"')
        return

    # chunk
    print('Chunking...')
    n = math.ceil(len(result) / 4)
    chunks = [ result[i:i+n] for i in range(0, len(result), n) ]

    # export
    print('Exporting...')
    for index, chunk in enumerate(chunks):
        chunk_path = path + f'-part{ index + 1 }.html'
        with open(chunk_path, 'w', encoding='utf-8', errors='ignore') as html_file:
            html_file.write(chunk)

    print(f'Successfully converted { len(chunks) } parts to => "converted/"')


def main():
    lang = input('Language (en|fr) =>')
    should_chunk = int(input('Chunk? (0|1)'))
    file_to_convert = input('File to convert (include extension) =>')
    file_path = f'./doc_files/{ file_to_convert }'
    converted_file = file_to_convert.split(".doc")[0]
    converted_path = f'./converted/{ converted_file }'

    # format
    with open(file_path, 'rb') as docx_file:
        result = mammoth.convert_to_html(
            docx_file,
            style_map=style_map
        ).value

    result = hc_sc_clean(result, lang)

    # export
    export(result, converted_path, should_chunk)

if __name__ == '__main__':
    main()
