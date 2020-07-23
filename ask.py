import xlrd
import string
import re
import jieba
import math
import tkinter
import gensim

jieba.load_userdict("术语表.txt")    #导入术语表
stopwords = [line.strip() for line in open('stopwords.txt', 'r', encoding='utf-8').readlines()]     #导入停用词表

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

def simister(str1,str2):    #计算相似度
    lenth1=len(str1)
    lenth2=len(str2)
    Union=list(set(str1).intersection(set(str2)))
    up=len(Union)
    down=(math.sqrt(lenth1)*math.sqrt(lenth2))
    return round(up/down,3)

def transfer(var1,var2):    #主程序函数
    workbook = xlrd.open_workbook(var1)
    booksheet = workbook.sheets()[0]                   #读取语料包
    rowNum = booksheet.nrows-1      #语料库长度
    Sstrs={}        #标准问字典
    Sdump=[]        #标准问集
    Scells=[]       #标准问切分

    for i in range(0,rowNum):       #装填标准问
        Sstrs.update({(str(booksheet.cell_value(i+1,1))): i})
        Sdump.append(str(booksheet.cell_value(i+1,1)))
    Sdics=list(Sstrs.keys())        #生成标准问列表

    for i in range(0,len(Sdics)):
        Scells.append(sentence2words(Sdics[i]))     #生成标准问分词
    Qcell=(sentence2words(var2))                    #生成普通问分词
            
    Sdictionary= gensim.corpora.Dictionary(Scells)
    num_features = len(Sdictionary.token2id)
    corpus = [Sdictionary.doc2bow(cell) for cell in Scells] #词包
    tfidf = gensim.models.TfidfModel(corpus)                #构建模型
    target = tfidf[corpus]

    listmid=[]  #储存回答信息，方便排序
    listend=[]
    tf_kw = tfidf[Sdictionary.doc2bow(Qcell)]
    similarities=gensim.similarities.SparseMatrixSimilarity(target, num_features).get_similarities(tf_kw)#相似度计算
    for i, simer in enumerate(similarities, 0):
        listmid.append([simer,Sdics[i]])

    listmid.sort(reverse=True)
    for j in range(0,5):
        listend.append([listmid[j][0],listmid[j][1]])
    return listend



window = tkinter.Tk()   #程序主入口
window.title('my window')
window.geometry('500x500')

def Start():            #按钮响应
    var1 = e1.get()
    var2=e2.get()
    s=transfer(var1,var2)
    for i in range(0,len(s)):
        t.insert('insert',i+1)
        t.insert('insert',':\t')
        t.insert('insert',s[i][1])
        t.insert('insert','\t')
        t.insert('insert',s[i][0])
        t.insert('insert','\n')

frm = tkinter.Frame(window)
frm.pack()
frm_1 = tkinter.Frame(frm, )
frm_1.pack()
frm_2 = tkinter.Frame(frm, )
frm_2.pack()

tkinter.Label(frm_1, text='请输入训练集文件名').pack(side='left')
e1 = tkinter.Entry(frm_1)
e1.pack(side='right')

tkinter.Label(frm_2, text='请输入问题').pack(side='left')
e2 = tkinter.Entry(frm_2)
e2.pack(side='right')

b1 = tkinter.Button(window, text='提问', command=Start)
b1.pack()
t = tkinter.Text(window, height=10)
t.pack()

window.mainloop()
