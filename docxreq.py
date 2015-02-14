try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile
import doorstop
from doorstop.common import DoorstopError

"""
Extract requirements from semi-structured MS Word (.docx) document
"""
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'

def _read_next(parg_iterator):
    next_parg = next(parg_iterator)
    return ''.join([x for x in next_parg.itertext()])

def _read_next_and_forward(parg_iterator):
    text = _read_next(parg_iterator)
    next(parg_iterator)
    return text

def _create(tree, path, value, parent):
    try:
        doc = tree.create_document(path=path, value=value, parent=parent)
    except DoorstopError as e:
        raise e
    else:
        return doc

def _find(tree, path, value, parent):
    try:
        doc = tree.find_document(value)
    except DoorstopError as e:
        raise e
    else:
        return doc


def process_document(repopath, tree, doctree, docfun=_create):
    paragraph = doctree.getiterator(PARA)
    input_reqs = []
    reqs = {}

    while True:
        try:
            text = _read_next(paragraph)
            if not text:
                continue

            if text == 'REQ_TYPE':
                prefix = _read_next_and_forward(paragraph)
                parent = _read_next(paragraph)
                treepath = '/'.join([repopath, prefix.lower()])

                try:
                    doc = docfun(tree,
                                 path=treepath,
                                 value=prefix,
                                 parent=parent)
                except Exception as e:
                    print(str(e))
                    break

                reqs = dict(zip([x.number for x in doc.items], 
                                [x.uid for x in doc.items]))

            elif text == 'REQ_NUM':
                for key in ['REQ_NUM', 'REQ_TEXT', 'REQ_RATIO', 'REQ_NOTE']:
                    field = _read_next_and_forward(paragraph)

                    if key == 'REQ_NUM':
                        num = int(field)
                        if num not in reqs:
                            # create
                            item = doc.add_item(num)
                            print('new requirement')
                        else:
                            # update
                            uid = prefix + '{:0>3d}'.format(num)
                            item = doc.find_item(uid)
                            print('update requirement')
                        input_reqs.append(num)

                    elif key == 'REQ_TEXT':
                        item.text = field
                    elif key == 'REQ_RATIO':
                        item.set(name='rationale', value=field)
                    elif key == 'REQ_NOTE':
                        item.set(name='note', value=field)

        except StopIteration:
            break

    if reqs:
        for key in (set(reqs) - set(input_reqs)):
            # delete
            uid = prefix + '{:0>3d}'.format(key)
            doc.remove_item(uid)
            print('delete requirement')

        print(doc.items)


repopath = './examples/reqs' # TODO
tree = doorstop.build(root=repopath)
print('On tree {}'.format(repopath))
print(tree)

while True:
    try:
        docpath = input('Add document: ')
        document = zipfile.ZipFile(docpath)
        xml_content = document.read('word/document.xml')
        document.close()
        doctree = XML(xml_content)
        process_document(repopath, tree, doctree, _create)
        more = str(input('Add more? '))
        if more == 'n':
            break
        else:
            continue

    except FileNotFoundError:
        print('File not found!')
        continue
