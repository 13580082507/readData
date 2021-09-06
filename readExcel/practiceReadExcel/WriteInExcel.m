function  M= WriteInExcel(x,y)
    x=rand(1)*10
    y=rand(1)*1
    
    s=xlswrite('zb.xls',x,1,'A2')
    s=xlswrite('zb.xls',y,1,'B2')
    
end

