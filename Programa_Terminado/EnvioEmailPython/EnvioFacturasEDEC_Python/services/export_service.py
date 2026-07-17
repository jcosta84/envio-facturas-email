from openpyxl import Workbook
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

def exportar_excel(relatorios,caminho):
    wb=Workbook(); ws=wb.active; ws.title='Relatórios'; ws.append(['Nome','E-mail','CIL','Status','Mensagem','Data'])
    for r in relatorios: ws.append([r.nome,r.email,r.cil,r.status,r.mensagem,r.data_envio.strftime('%d/%m/%Y %H:%M')])
    ws.freeze_panes='A2'; wb.save(caminho)

def exportar_pdf(relatorios,caminho):
    doc=SimpleDocTemplate(caminho,pagesize=landscape(A4),leftMargin=20,rightMargin=20,topMargin=25,bottomMargin=25)
    dados=[['Nome','E-mail','CIL','Status','Mensagem','Data']]
    for r in relatorios: dados.append([r.nome,r.email,r.cil,r.status,r.mensagem[:60],r.data_envio.strftime('%d/%m/%Y %H:%M')])
    tabela=Table(dados,repeatRows=1,colWidths=[100,145,70,65,220,95]); tabela.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('GRID',(0,0),(-1,-1),.4,colors.grey),('FONTSIZE',(0,0),(-1,-1),7),('VALIGN',(0,0),(-1,-1),'TOP')]))
    doc.build([Paragraph('Relatório de Envios',getSampleStyleSheet()['Title']),tabela])
