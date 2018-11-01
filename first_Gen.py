class xlsOperator():
    
    def __init__(self, df, name):
        
        
        self.writer = pd.ExcelWriter(name + '.xlsx', engine = 'xlsxwriter')
        self.df = df
        (self.df).to_excel(self.writer, sheet_name = 'Custom')
        self.workbook = (self.writer).book
        self.worksheet = (self.writer).sheets['Custom']
        
    
        
    def applyFormulas_Column(self,formula_String, column, applyColumn = False):
        
        
# =============================================================================
#          column : Result column (char)
#          
#          formula_String : Raw formula (string)
#          
#          applyColumn : In next versions method can handle multiple colum formulas (boolean)
#              
#           !  ! ! Applies formulas on one column  ! ! !
# =============================================================================
         
         
   
        if applyColumn == True:
            
            formula_String = list(formula_String)
            
            columns_List = list()
            
            opt_Sign_List = list()
            
            for opt in formula_String:
                
                # Separate operant and operator
                
                if  str.isalnum(opt) :
                    
                    columns_List.append(opt)
                
                elif not str.isalpha(opt) :  
                    
                    opt_Sign_List.append(opt)
            
            
            formula = self.formula_Constructor(columns_List, opt_Sign_List)
            
            for i in (self.df).index + 2:
                
                
                # Replace spaces to column numbers
                temp_Formula = formula.replace(' ', str(i)) 
                self.worksheet.write_formula(column + str(i), temp_Formula)
                
                    
       
                
            
        
        
        return 0
    

    
    def interval_Parser_Border(self, interval):
        
        # Box selecter for future work
        
        for letter in interval:
            
            if  str.isalpha(letter) :
                
                interval = interval.replace(letter,"")
        
               
        parsed_List = interval.split(':')
                
        
        return parsed_List[0]; parsed_List[1]
        
    def applyVisual_Header(header_Row, interval):
        
        # Apply simple header design
        
        
        pass
    
    def formula_Constructor(self,columns, operator):
        
# =============================================================================
#         Generates constructible formula
#         
#         Add black spaces after Columns for column nublers
# =============================================================================
        
    
        
        formula_List = list()
        
        columns_Temp = columns.copy() # For tracing of the list we coied it temporary variable
        operator_Temp = operator.copy() # list_1 = list_2 assignment doent work, because of the reference copy
        
        columns_Temp.reverse() # Usage of pop method causes to use reverse form of the list
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
