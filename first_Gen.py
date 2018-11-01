class xlsOperator():
    
    def __init__(self, df, name):
        
        
        self.writer = pd.ExcelWriter(name + '.xlsx', engine = 'xlsxwriter')
        self.df = df
        (self.df).to_excel(self.writer, sheet_name = 'Custom')
        self.workbook = (self.writer).book
        self.worksheet = (self.writer).sheets['Custom']
        
    
        
    def applyFormulas_Column(self,formula_String, column, applyColumn = False):
        
        
            
   
        if applyColumn == True:
            
            formula_String = list(formula_String)
            
            columns_List = list()
            
            opt_Sign_List = list()
            
            for opt in formula_String:
                
                if  str.isalnum(opt) :
                    
                    columns_List.append(opt)
                
                elif not str.isalpha(opt) :  
                    
                    opt_Sign_List.append(opt)
            
            
            formula = self.formula_Constructor(columns_List, opt_Sign_List)
            
            for i in (self.df).index + 2:
                
                temp_Formula = formula.replace(' ', str(i)) 
                self.worksheet.write_formula(column + str(i), temp_Formula)
                
                    
       
                
            
        
        
        return 0
    

    
    def interval_Parser_Border(self, interval):
        
        
        for letter in interval:
            
            if  str.isalpha(letter) :
                
                interval = interval.replace(letter,"")
        
               
        parsed_List = interval.split(':')
                
        
        return parsed_List[0]; parsed_List[1]
        
    def applyVisual_Header(header_Row, interval):
        
        pass
    
    def formula_Constructor(self,columns, operator):
        
        formula_List = list()
        
        columns_Temp = columns.copy()
        operator_Temp = operator.copy()
        
        columns_Temp.reverse()
        operator_Temp.reverse()
        

        
        for i in columns:
            
            if str.isalpha(i):
                
                formula_List.append(columns_Temp.pop())
                formula_List.append(" ")
                
                
                try:
                    formula_List.append(operator_Temp.pop())
                
                except:
                    
                    continue
                
            else:
                
                formula_List.append(columns_Temp.pop())
                
            
                
        
        
        return ''.join(formula_List)
