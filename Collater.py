'''

Developed by: Zaahier Adams
https://github.com/ZaahierAdams

'''

from glob import glob
import pandas as pd
import os

class Excel_Collater():
 
    success = False
    
    #def main():
    def __init__(self,index):
        
        # Get collated file
        self.df = Excel_Collater.Collate(index)
        
        
        # Save colalted file 
        try:
            os.mkdir('Collate')
        except:
            pass
        
        cDir= os.getcwd()
    
        #df.to_csv(cDir+'\Collate\Collate.csv', index = True, sep=';')
        self.df.to_excel(cDir+'\Collate\Collate.xlsx')
        print(self.df)
        
        
        # Result of collation
        Excel_Collater.Success()

    @classmethod   
    def Success(cls):
        cls.success = True
          
        
    def NGT(fname, IndexName):
        '''
        Find length of Index
        '''
        
        df = (pd.read_excel(fname)).apply(lambda x: x.nunique())
        index = df[IndexName]
        
        return index
    
    
    def Duplicate_Columns(df):
        '''
        Get a list of duplicate columns.
        '''  
        ## DCN = Duplicate Column Names
        DCN = set()
        
        for x in range(df.shape[1]):
            col = df.iloc[:, x]
            for y in range(x + 1, df.shape[1]):
                otherCol = df.iloc[:, y]
                if col.equals(otherCol):
                    DCN.add(df.columns.values[y])
        return list(DCN)
    
    def Dup_Col_Names(Index, df,c1,c2):
        # This is a redundant check. Use -> Duplicate_Columns(df)
        # Detect identical column(s) in Dataframe 2
        # Detects by column name, not by column content
        remove = []
        for a in c2:
            for b in c1:
                if a == b:
                    remove.append(a)
                else:
                    pass
                
        # Remove redundant column(s)
        # Except for index column(s)    
        remove.remove(Index)
        for e in remove:
            df = df.drop(columns = e)
    
        return df
    
    
    def Collate(Index):
        '''
        Collates xlsx data 
        '''
        
        count = 0
        
        for fname in glob('*.xlsx'):
            
            # Set Dataframe 1 (df1)
            # All subsequent dataframes will be merged to df1 
            if count ==0:
                
                # look at columns
                c1 = list(pd.read_excel(fname).columns)
                
                # data frame    
                df1 = pd.DataFrame(pd.read_excel(fname, names=pd.read_excel(fname), index_col = Index))
                
                # NGT
                ngt_1 = Excel_Collater.NGT(fname, Index)
        
            else:
                
                # Set Dataframe 2
                c2 = list(pd.read_excel(fname).columns)
                df2 = pd.read_excel(fname, names=c2, index_col = Index)
                df2 = pd.DataFrame(df2)
                
                
                ngt_2 = Excel_Collater.NGT(fname, Index)
                
                # If index in df2 is identical to df1
                if ngt_2 == ngt_1:
                    df2 = Excel_Collater.Dup_Col_Names(Index, df2,c1,c2)
                    
                    
                    # Merge current df2 to df1
                    df1 = pd.merge(df1, df2, on=Index)
                
                
                # If index in df2 is NOT identical to df1
                # !!!Fix for duplicate columns in append files!!!
                else:                    
                    df1 = df1.append(df2,sort=True)
                                
            count += 1
               
        return df1
    
    

     
#    if __name__ == '__main__':
#        main()
#                
