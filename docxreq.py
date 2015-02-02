try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile
import doorstop

"""
Extract requirements from semi-structured MS Word (.docx) document
"""
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'

class Document(object):
    def __init__(self, docpath, repopath):
        document = zipfile.ZipFile(docpath)
        xml_content = document.read('word/document.xml')
        document.close()
        self.doctree = XML(xml_content)
        self.repopath = repopath
        self.tree = doorstop.build(root=repopath)

    def _read_next(self, parg_iterator):
        next_parg = next(parg_iterator)
        return ''.join([x for x in next_parg.itertext()])

    def _read_next_and_forward(self, parg_iterator):
        text = self._read_next(parg_iterator)
        next(parg_iterator)
        return text

    def sync_requirements(self):
        paragraph = self.doctree.getiterator(PARA)
        input_reqs = []

        while True:
            try:
                text = self._read_next(paragraph)
                if not text:
                    continue

                if text == 'REQ_TYPE':
                    prefix = self._read_next_and_forward(paragraph)
                    parent = self._read_next(paragraph)
                    treepath = [self.repopath] if parent == '' else [self.repopath, parent]

                    try:
                        doc = self.tree.create_document(path='/'.join(treepath),
                                                        value=prefix,
                                                        parent=parent)
                    except doorstop.common.DoorstopError:
                        doc = self.tree.find_document(prefix)

                    reqs = dict(zip([x.number for x in doc.items], 
                                    [x.uid for x in doc.items]))

                elif text == 'REQ_NUM':
                    for key in ['REQ_NUM', 'REQ_TEXT', 'REQ_RATIO', 'REQ_NOTE']:
                        field = self._read_next_and_forward(paragraph)

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

        for key in (set(reqs) - set(input_reqs)):
            # delete
            uid = prefix + '{:0>3d}'.format(key)
            doc.remove_item(uid)
            print('delete requirement')

        print(doc.items)
