from win32com.client import Dispatch

word = Dispatch('Word.Application')

word.Visible = 0        
word.DisplayAlerts = 0  
import os
from glob import glob
with open('report.csv', 'w') as writer:
    for doc in glob('docs/*'):
        path = os.path.abspath(doc)
        doc = word.Documents.Open(FileName=path, Encoding='gb2312')

        for t in doc.Tables:
            for row in range(2,t.Rows.Count+1):
                for col in range(1, t.Columns.Count+1):
                    writer.write(t.Cell(Row=row, Column=col).Range.Text.replace('\r','').replace('\x07','')+',')
                writer.write('\n')

        doc.Close()
        print('Done '+ path)

word.Quit