
"""
   conversion Survey, Excle vers PDF
   Fait par Jean-Paul Varga  - le 19/11/2017

    


    """




import xlrd
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.units import inch


logo = ".\\logo.png"
logo2 = ".\\mi.jpg"
bn = ".\\pi.jpg"

f='.\\fepcc.xlsx'

q1="Identité de l'enfant"
q2="Classe"
q3="Identité parent"
q4="Rien à signaler"
q5="A propos de votre enfant : - Quelles sont les informations personnelles ou familiales que vous souhaitez nous communiquer ?"
q6="Votre enfant rencontre-t-il des difficultés scolaires pour ce trimestre ? Pourquoi ?"
q7="Sentez-vous votre enfant intégré et épanoui au sein du lycée/collège/annexe ? Dans sa classe ? Pourquoi ?"
q8="Quels sont les points positifs de ce trimestre ?"
q9="N'hésitez pas à prendre rendez-vous avec le(s) professeur(s) concerné(s) via le carnet de correspondance de votre enfant."
q10="Souhaitez-vous nous faire part d'autres informations ?"
r2=""
nc = 0
doc="vide"
init=0



def addPageNumber(canvas, doc):
    """
    Add the page number
    """
    page_num = canvas.getPageNumber()
    text = "Page - %s - " % page_num
    canvas.drawRightString(200 * mm, 20 * mm, text)


# ----------------------------------------------------------------------
def createpdf():
    """
   pour plus tard..
    """
    doc = SimpleDocTemplate(f1, pagesize=letter,
                            rightMargin=72, leftMargin=72,
                            topMargin=72, bottomMargin=18)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Justify', alignment=TA_JUSTIFY))

    Story = []







workbook = xlrd.open_workbook(f)
sheet = workbook.sheet_by_index(0)
for rowx in range(sheet.nrows):
    cols = sheet.row_values(rowx)
    r1=cols[9]+" "+cols[10] # nom + prenom de l'enfant
    cc = r2
    r2=cols[11] # classe rectifié par le parent


    r3=cols[13]+" "+cols[14] # nom + prenom parent
    r4=cols[16] # Si vous n'avez rien à signaler, cochez R.A.S. et "Envoyer" le formulaire en bas de page. [Rien à signaler]
    r5=cols[17] # A propos de votre enfant :    	- Quelles sont les informations personnelles ou familiales que vous souhaitez nous communiquer ?
    r6 = cols[18] # - Votre enfant rencontre-t-il des difficultés scolaires pour ce  trimestre ? Pourquoi ?
    r7 = cols[19]
    r8 = cols[20]
    r9 = cols[21]
    r10 = cols[22]
    r11 = cols[6] # réponse



    print (r2)
    print (r3)
    print (r4)
    print (r5)
    print (len(r5))




    if nc == 1:
        f1 = '.\\res\\'+ r2 + '-fepcc-rapport-T1.pdf'
        doc = SimpleDocTemplate(f1, pagesize=letter,
                                rightMargin=72, leftMargin=72,
                                topMargin=72, bottomMargin=18)
        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name='Justify', alignment=TA_JUSTIFY))
        Story = []
        im = Image(logo, 5 * inch, 1 * inch)
        Story.append(im)


        Story.append(Spacer(1, 20))
        ptext = '<font size=15 color=blue  >          Préparation conseil de classe  : APEL Marcq Institution  </font>'
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(15, 22))



        ptext = '<font size=12>Ensemble des commentaires pour les élèves de la classe de %s </font>' % r2
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 22))

        ptext = '<font size=18 >    </font>'
        Story.append(Paragraph(ptext, styles["Normal"]))

        im = Image(logo2, 6 * inch, 2 * inch)
        Story.append(im)

        ptext = '<font size=18 >    </font>'
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 20))

        ptext = '<font size=14 >Trimestre 1    Année 2018-2019  (c) JPV </font>'
        Story.append(Paragraph(ptext, styles["Normal"]))

        ptext = '<font size=14 >Confidentiel - à destination uniquement des parents correspondants  </font>'
        Story.append(Paragraph(ptext, styles["Normal"]))

        Story.append(Spacer(11, 21))

        Story.append(PageBreak())
        nc = 0


    if r2 != cc:
        if doc != "vide":
            doc.build(Story, onFirstPage=addPageNumber, onLaterPages=addPageNumber)
        nc = 1
    if doc != "vide":
        im = Image(logo, 5 * inch, 1 * inch)
        Story.append(im)
        Story.append(Spacer(5, 12))
        ptext = '<font size=14  >Nom élève : %s</font>' % r1
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))
        ptext = '<font size=12>Classe : %s</font>' % r2
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))
        ptext = '<font size=12>Parents : %s </font>' % r3
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 22))
        ptext = '<font size=12 color=red >Tout va bien pour l\'enfant (rien à signaler) : %s</font>' % r4
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))

        ptext = '<font size=12 color=green >Question 1 : A propos de votre enfant : - Quelles sont les informations personnelles ou familiales que vous souhaitez nous communiquer ? </font>'
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))
        ptext = '<font size=12>%s </font>' % r5
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))

        ptext = '<font size=12 color=green >Question 2 : Votre enfant rencontre-t-il des difficultés scolaires pour ce trimestre ? Pourquoi ? </font>'
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))
        ptext = '<font size=12>%s </font>' % r6
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))

        ptext = '<font size=12 color=blue  >Question 3 : Sentez-vous votre enfant intégré et épanoui au sein du lycée/collège/annexe ? Dans sa classe ? Pourquoi </font>'
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))
        ptext = '<font size=12>%s </font>' % r7
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 12))




        ptext = '<font size=12 color=green >Question 5 : Quels sont les points positifs de ce trimestre ? </font>'
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))
        ptext = '<font size=12>%s </font>' % r8
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))

        ptext = '<font size=12 color=green >Question 6 : N\' hesitez pas à prendre rendez-vous avec le(s) professeur(s) concerné(s) via le carnet de correspondance de votre enfant </font>'
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))
        ptext = '<font size=12>%s </font>' % r9
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))

        ptext = '<font size=12 color=green >Question 7 : Souhaitez-vous nous faire part d\'autres informations ? </font>'
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))

        ptext = '<font size=12>%s </font>' % r10
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))

        im = Image(bn , 4 * inch, 2 * inch)
        Story.append(im)
        ptext = '<font size=8 color=pink > Post-it pour apporter des remarques sur %s </font>' % r1
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))
        ptext = '<font size=6 color=blue  > réponse au questionnaire  %s </font>' % r11
        Story.append(Paragraph(ptext, styles["Normal"]))
        Story.append(Spacer(1, 10))

        Story.append(PageBreak())




