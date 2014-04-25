# -*- coding: utf-8 -*- 
import xdrlib,sys,os,xlrd,string


class os_:
    def get_dir_fileinfo(self,path,*fileext):
        results =[]
        for root, dirs, files in os.walk(path ):
            for fn in files:
                filename=sufix = os.path.splitext(fn)[0]
                sufix = os.path.splitext(fn)[1] 
                if(len(fileext)==0 or sufix in fileext):                
                    fileinfo=(root,filename,sufix)
                    results.append(fileinfo)
        return results
    def get_aciistr(self,cell_valuetmp):
        #try GB2312,UTF8,gbk
        try:
            #cell_value='"'+cell_valuetmp.encode('gb2312').replace('"',"'")+'"'  
            cell_value=cell_valuetmp.encode('gb2312')  
        except:
            try:
                cell_value = str(cell_valuetmp.encode('gbk'))
            except:
                cell_value = str(cell_valuetmp.encode('utf8'))
        return cell_value
            
    

class excel_:
    #print(sys.getdefaultencoding()) 
    def open_excel(self,file):
        try:
            data = xlrd.open_workbook(file)
            return data
        except Exception,e:
            print str(e)
            
    def xls_2_csv_s(self,xlsfile,csvfile,isforce=0,sh_index=0):
        if(not os.path.exists(xlsfile)):
            return ('Error:XLS File does not exist!')
        if(os.path.exists(csvfile)):
            if(isforce==0):            
                return ('OK:CSV File already exists, Skip!')
            else:
                os.remove(csvfile)            
        iserror=False
        errormess=u''
        data = self.open_excel(xlsfile)
        table = data.sheets()[sh_index]
        nrows_num = table.nrows 
        ncols_num = table.ncols 
        fp = open(csvfile,'a')
        for nrows in range(nrows_num):
            row_value=''    
            for ncols in range(ncols_num):        
                cell_valuetmp = table.cell(nrows,ncols).value
                if isinstance(cell_valuetmp,unicode):
                    try:
                        cell_value=os_().get_aciistr(cell_valuetmp)
                        #cell_value='"'+cell_value.replace('"',"'")+'"'
                    except Exception,ex:
                        print(cell_valuetmp)
                        errormess=('ex_'+str(ex))
                        iserror=True
                        return errormess
                else:
                    try:                    
                        cell_value = str(cell_valuetmp)                         
                    except Exception,ec:
                        print(cell_valuetmp)
                        errormess='ec_'+str(ec)
                        iserror=True 
                        return errormess
                if ',' in cell_value:
                            cell_value='"%s"' % (cell_value.replace('"',"'")) 
                row_value="%s%s," % (row_value,cell_value)
            if(len(row_value)>1):
                row_value=row_value[:-1]+'\n'
                fp.write(row_value)   
        fp.close()
        if(iserror):
            #remove csvFile if ERROR
            if(os.path.exists(csvfile)):
                os.remove(csvfile)
            return 'Fail:'+errormess
        return 'OK:'+str(nrows_num)+'Rows,'+str(ncols_num)+'Column'

    
                    
    def xls_2_csv(self,xlsdir,isforce=0):
        try:
            files=os_().get_dir_fileinfo(xlsdir,'.xls','.xlsx')
            for dirname,filename,sufix in files:
                xlsfile=os.path.join(dirname, filename+sufix)
                csvfile=os.path.join(dirname, filename+'.csv')
                print(xlsfile+' --> '+csvfile)
                try:
                    print ('    '+self.xls_2_csv_s(xlsfile,csvfile,isforce))
                except Exception,ed:        
                    print 'ed_'+str(ed)        
        except Exception,e:        
            print str(e)
       
if __name__=="__main__":    
    #Method1:single Excel file
    #转换单个文件（EXCEL文件名，CSV文件名，是否覆盖执行）
    #excel_().xls_2_csv_s(r'D:\pri\GitHub\Excel2Csv\testfile\test.xls',r'D:\pri\GitHub\Excel2Csv\testfile\test.csv',0)
    #Method2:All Excel files in a directory
    #转换文件夹下的文件（目录名，是否覆盖执行）
    excel_().xls_2_csv(r'D:\pri\GitHub\Excel2Csv\testfile',1)
    

