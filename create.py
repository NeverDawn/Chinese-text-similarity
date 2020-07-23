import xlrd
import string
import re
import jieba
import math
import tkinter
import gensim

jieba.load_userdict("术语表.txt")    #导入术语表
stopwords = [line.strip() for line in open('stopwords.txt', 'r', encoding='utf-8').readlines()]

syn=[]
with open('synonyms.txt','r') as file:
    for line in file:
        synim=[]
        for word in line.strip('\n').split('\t'):
            synim.append(word)
        syn.append(synim)

def sentence2words(strin):  #分词预处理
    outstr=[]
    mid=list(jieba.cut(strin, cut_all=False))
    for word in mid:  
        if word not in stopwords:  
            outstr.append(word)
    for i in range(0,len(outstr)):
        for j in range(0,len(syn)):
            if outstr[i] in syn[j]:
                outstr[i]=syn[j][0]
    return outstr


def transfer(var):          #主程序函数
    workbook = xlrd.open_workbook(var)
    booksheet = workbook.sheets()[0]                   #读取语料包
    rowNum = booksheet.nrows-1                          #语料库长度
    Sstrs={}        #标准问字典
    Sdump=[]        #标准问集
    Qdump=[]        #普通问集
    Scells=[]       #普通问切分
    Qcells=[]       #标准问切分


    for i in range(0,rowNum):       #装填问题
        Sstrs.update({(str(booksheet.cell_value(i+1,1))): i})
        Qdump.append(str(booksheet.cell_value(i+1,0)))
        Sdump.append(str(booksheet.cell_value(i+1,1)))
    Sdics=list(Sstrs.keys())    #标准问列表
    Qdics=Qdump                 #普通问列表



    for i in range(0,len(Sdics)):
        Scells.append(sentence2words(Sdics[i])) #生成标准问分词
    for i in range(0,len(Qdics)):
        Qcells.append(sentence2words(Qdics[i])) #生成普通问分词

        
    Sdictionary= gensim.corpora.Dictionary(Scells)
    num_features = len(Sdictionary.token2id)
    corpus = [Sdictionary.doc2bow(cell) for cell in Scells] #词包
    tfidf = gensim.models.TfidfModel(corpus)                #构建模型
    target = tfidf[corpus]

    
    f1 = open("answer.txt","w")      #top5及结果存放地址
    number=0                #相似度0.8+计数器
    outer=0                 #正确回答计数器

    for j in range(0,len(Qcells)):
        listmid=[]
        tf_kw = tfidf[Sdictionary.doc2bow(Qcells[j])]
        similarities=gensim.similarities.SparseMatrixSimilarity(target, num_features).get_similarities(tf_kw)
        for i, simer in enumerate(similarities, 0):
            liststart=[simer,Qdics[j],Sdump[j],Sdics[i]]
            listmid.append(liststart)


        listmid.sort(reverse=True)

        for item in range(0,5):
            f1.write(str(listmid[item][0])+"\n")
            f1.write(listmid[item][1]+"\n") #普通问
            f1.write(listmid[item][2]+"\n") #正确
            f1.write(listmid[item][3]+"\n") #匹配
            f1.write("************"+"\n")

        f1.write("-------------------------"+"\n")
        if(listmid[0][0]>=0.8):
            number=number+1
            if(listmid[0][2]==listmid[0][3]):
                    outer=outer+1

    p=outer/number
    r=outer/rowNum
    f1.write("相似度0.8以上："+str(number)+"\n")
    f1.write("正确数："+str(outer)+"\n")
    f1.write("P:"+str(p)+"\n")
    f1.write("R:"+str(r)+"\n")
    f1.write("F1:"+str(2*p*r/(p+r))+"\n")
    f1.close()
    print("ok")


window = tkinter.Tk() #程序主入口
window.title('my window')
window.geometry('500x500')



def Start():        #按钮响应
    var = e1.get()
    transfer(var)



frm = tkinter.Frame(window)
frm.pack()
frm_1 = tkinter.Frame(frm, )
frm_1.pack()


tkinter.Label(frm_1, text='请输入训练集文件名').pack(side='left')
e1 = tkinter.Entry(frm_1)
e1.pack(side='right')
b1 = tkinter.Button(window, text='生成', command=Start)
b1.pack()

window.mainloop()
