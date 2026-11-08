from pathlib import Path
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as col

FIRST=7
CAP={'SETUP':20,'BELANJA':1500,'KAS USAHA':1500,'NOTA':1500,'ITEM':6000,'STOK':20}

def rng(sheet,c):return f"'{sheet}'!${c}$7:${c}${6+CAP[sheet]}"
def look(value,sheet,result,default='""'):
    return f'IFERROR(INDEX({rng(sheet,result)},MATCH({value},{rng(sheet,"A")},0)),{default})'
def checks(empty,conditions,default='"OK"'):
    s=default
    for condition,message in reversed(conditions):s=f'IF({condition},"{message}",{s})'
    return f'IF({empty}="","",{s})'

class Layout:
    def __init__(self,path):
        self.w=xlsxwriter.Workbook(str(path));self.w.set_calc_mode('auto');self.s={};self.f={}
        self.w.set_properties({'title':'Chicken Al-Haqi | POS Sederhana','author':'Social Project BSI','comments':'MoM 23 Oktober 2026; transaksi operasional kosong; VBA terpisah.'})
        common={'font_name':'Calibri','font_size':11,'font_color':'#404040'}
        styles={
          'title':{'font_size':19,'bold':True,'bg_color':'#902A20','font_color':'white','valign':'vcenter'},
          'sub':{'font_size':10,'font_color':'#808080','text_wrap':True,'valign':'vcenter'},
          'nav':{'font_color':'white','bg_color':'#C96B1D','bold':True,'font_size':9,'align':'center','valign':'vcenter','text_wrap':True},
          'head':{'bg_color':'#A63B25','font_color':'white','bold':True,'text_wrap':True,'valign':'vcenter'},
          'input':{'bg_color':'#FFFFFF','font_color':'#902A20','bottom':1,'bottom_color':'#D9D9D9'},
          'calc':{'bg_color':'#F2F2F2'},
          'body':{'text_wrap':True,'valign':'top'},
          'label':{'bold':True,'font_color':'#808080'},
          'note':{'text_wrap':True,'font_size':10,'font_color':'#808080','valign':'top'},
          'warn':{'text_wrap':True,'bg_color':'#FFF1CE','font_color':'#7A570B'},
          'bad':{'bg_color':'#FBE4DF','font_color':'#982A20'},
          'good':{'bg_color':'#E4F1E7','font_color':'#22663B'},
          'kpi_label':{'bg_color':'#F9E8D2','font_color':'#902A20','font_size':10},
          'kpi':{'bg_color':'#FBF3E7','font_color':'#902A20','bold':True,'font_size':23,'num_format':'#,##0;[Red](#,##0);"0"'},
        }
        for name,p in styles.items():self.f[name]=self.w.add_format({**common,**p})
        self.f['ref_time']=self.w.add_format({'font_name':'Calibri','font_size':11,'bold':True,'font_color':'#902A20','bg_color':'white','align':'center','valign':'vcenter','num_format':'hh:mm'})
        for name,num in [('money','#,##0;[Red](#,##0);"–"'),('qty','0.###;[Red](0.###);"–"'),('date','dd mmm yyyy'),('time','hh:mm'),('pct','0.0%')]:
            for kind in ['calc','input']:
                self.f[name if kind=='calc' else 'in_'+name]=self.w.add_format({**common,**styles[kind],'num_format':num})
    def base(self,name,subtitle,width=12):
        s=self.w.add_worksheet(name);self.s[name]=s;s.hide_gridlines(2);s.set_default_row(20)
        s.set_column(0,width-1,17);s.set_tab_color('#902A20' if name in ['MULAI','DASHBOARD','KASIR','PRODUKSI'] else '#808080')
        s.merge_range(0,0,0,width-1,'CHICKEN AL-HAQI' if name=='MULAI' else name,self.f['title']);s.set_row(0,34)
        s.merge_range(1,0,1,width-1,subtitle,self.f['sub']);s.set_row(1,28)
        for i,n in enumerate(['MULAI','KASIR','HARIAN','NOTA','STOK','TUTUP KASIR']):s.write_url(3,i*2,f"internal:'{n}'!A1",self.f['nav'],n.title())
        s.freeze_panes(6,2);s.set_zoom(80);s.set_landscape();s.set_paper(9);s.fit_to_pages(1,0)
        s.set_margins(.25,.25,.4,.4);s.set_footer('&LChicken Al-Haqi • Internal&R&P / &N');s.repeat_rows(0,5)
        return s
    def table(self,s,name,headers,capacity,formulas=None,inputs='',formats=None,first=7):
        formats=formats or {};formulas=formulas or {}
        columns=[{'header':h,'header_format':self.f['head'],'format':self.f[formats.get(col(i),'input' if col(i) in inputs else 'calc')]} for i,h in enumerate(headers)]
        s.add_table(first-2,0,first-2+capacity,len(headers)-1,{'name':name,'style':'Table Style Light 9','columns':columns});s.set_row(first-2,36)
        for r in range(first,first+capacity):
            for c,fn in formulas.items():s.write_formula(c+str(r),'='+fn(r),self.f[formats.get(c,'calc')],'')
        s.print_area(0,0,min(60,5+capacity),len(headers)-1)
    def validate(self,s,where,source):s.data_validation(where,{'validate':'list','source':source,'error_type':'stop','error_message':'Pilih kode atau opsi dari daftar.'})
    def audit(self,s,column,last):
        s.set_column(column+':'+column,25)
        s.conditional_format(f'{column}7:{column}{last}',{'type':'formula','criteria':f'AND({column}7<>"",{column}7<>"OK")','format':self.f['bad']})
    def formula(self,s,cell,formula,kind='calc',value=''):s.write_formula(cell,'='+formula,self.f[kind],value)
    def close(self,order):
        self.w.worksheets_objs=[self.s[n] for n in order]
        for i,s in enumerate(self.w.worksheets_objs):s.index=i
        self.s['KASIR'].activate();self.w.close()
