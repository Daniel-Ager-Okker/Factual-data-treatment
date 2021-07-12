import os
files=os.listdir()
os.mkdir('обработанные')
print(os.getcwd())
for file in files:
    A=[]
    if file.endswith('.dev'):
        fin=open(file,'r')
        for line in fin:
            if'#' in line or 'MD' in line:
                pass
            else:
                A.append(line.split()[1]+' '+line.split()[2]+' '+str(float(line.split()[3])*(-1))+' '+line.split()[0]+'\n')
        fin.close()
        fout=open(f'{os.getcwd()}/обработанные/{file}','w')
        fout.write(f"WELLTRACK {file[:file.rindex('.')]}\n")
        for a in A:
            fout.write(a)
        fout.close()