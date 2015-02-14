"""
Extract requirements from semi-structured MS Word (.docx) document
"""

try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile
import doorstop
import readline
import argparse
from doorstop.common import DoorstopError


'''
.docx parser utility functions
'''
def get_xml_tree(docxpath):
    document = zipfile.ZipFile(docxpath)
    xml_content = document.read('word/document.xml')
    document.close()
    return XML(xml_content)

'''
Doorstop wrapper functions
'''
def _create(tree, path, value, parent):
    try:
        doc = tree.create_document(path=path, value=value, parent=parent)
    except DoorstopError as exc:
        raise exc
    else:
        return doc

def _find(tree, path, value, parent):
    try:
        doc = tree.find_document(value)
    except DoorstopError as exc:
        raise exc
    else:
        return doc

'''
Processing functions: read .docx and entry to Doorstop
'''
def _read_next(parg_iterator):
    next_parg = next(parg_iterator)
    return ''.join([x for x in next_parg.itertext()])

def _read_next_and_forward(parg_iterator):
    text = _read_next(parg_iterator)
    next(parg_iterator)
    return text

def process_document(repopath, tree, doctree, docfun=_create):
    word_namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    para = word_namespace + 'p'
    paragraph = doctree.getiterator(para)
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
                except Exception as exc:
                    print(str(exc))
                    break

                reqs = dict(zip([x.number for x in doc.items],
                                [x.uid for x in doc.items]))

            elif text == 'REQ_NUM':
                for key in ['REQ_NUM', 'REQ_LINKS', 'REQ_TEXT', 'REQ_RATIO', 'REQ_NOTE']:
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

                    elif key == 'REQ_LINKS':
                        item.links = [] # first remove all links
                        links = [x.strip() for x in field.split(',') if x != '']
                        for link in links:
                            item.link(link)

                    elif key == 'REQ_TEXT':
                        item.text = field
                    elif key == 'REQ_RATIO':
                        item.set(name='rationale', value=field)
                    elif key == 'REQ_NOTE':
                        item.set(name='note', value=field)

        except StopIteration:
            break

    if reqs:
        for key in set(reqs) - set(input_reqs):
            # delete
            uid = prefix + '{:0>3d}'.format(key)
            doc.remove_item(uid)
            print('delete requirement')


if __name__ == '__main__':
    PARSER = argparse.ArgumentParser(description='Parse .docx documents into Doorstop files')
    PARSER.add_argument('repopath', type=str, help='path to the requirement tree')
    ARGS = PARSER.parse_args()

    readline.parse_and_bind('tab: complete')
    readline.parse_and_bind('set editing-mode vi')

    REQTREE = doorstop.build(root=ARGS.repopath)
    print('On tree {}'.format(ARGS.repopath))
    print(REQTREE)

    while True:
        try:
            print('1. Add document')
            print('2. Update document')
            print('3. Analyze requirement tree')
            print('4. Quit')
            SEL = int(input('> '))

            if SEL == 1:
                DOCPATH = input('Document path: ')
                DOCTREE = get_xml_tree(DOCPATH)
                process_document(ARGS.repopath, REQTREE, DOCTREE, _create)

            elif SEL == 2:
                DOCPATH = input('Document path: ')
                DOCTREE = get_xml_tree(DOCPATH)
                process_document(ARGS.repopath, REQTREE, DOCTREE, _find)

            elif SEL == 3:
                for issue in REQTREE.issues:
                    print(issue)

            elif SEL == 4:
                break

            else:
                continue

        except KeyboardInterrupt:
            break

        except Exception as exc:
            print(str(exc))
            continue
