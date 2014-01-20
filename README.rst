==============
docxgen README
==============

docxgen is a python library to generate Microsoft Office Word 2007 Document.
Without knowing the details of the WordprocessingML, you may create a 
beautiful Word Document, as:

::

    from docxgen import Document, title
    doc = Document()
    doc.update(
        title='Moby Dick',
        creator='Herman Melville',
        created=datetime(...)
    )

    body = doc.body
    body.append(title('Moby Dick'))
    body.append(subtitle('Moby Dick'))
    body.append(h1('Chapter 1'))
    body.append(paragraph([run('Call me Ishmael. Some years ago ...')]))

    doc.save('/tmp/moby-dick.docx')


This work is proudly sponsored by `SurveyMonkey Inc`, and licensed in MIT license.

.. _SurveyMonkey Inc: https://www.surveymonkey.com
.. _MIT license: http://opensource.org/licenses/MIT
