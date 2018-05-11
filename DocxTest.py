import vsts_datapack.DatapackDocument as dd

if __name__ == '__main__':

    my_doc = dd.DatapackDocument()
    my_doc.add_paragraph('This is a para ')
    my_doc.add_text('adding some text!')
    my_doc.add_heading('Heading, this is one', 1)
    my_doc.add_text('adding some text!')
    my_doc.add_paragraph('YO NEW PARA')
    my_doc.page_break()
    my_doc.add_text('adding NEWEST TEXT')
    my_doc.add_paragraph('YO NEW PARA 222')
    my_doc.add_text('TEXT #ETS F#RWEF DS')
    my_doc.write('C:\\Users\\v-maubut\\PycharmProjects\\vsts_datapack\\Testdocx.docx')



