from io import BytesIO
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER

from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.styles import ParagraphStyle
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.textlabels import Label

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Image,
    PageBreak,
    Table,
    TableStyle
)
def generate_executive_pdf(df):

    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        rightMargin=35,
        leftMargin=35,
        topMargin=40,
        bottomMargin=35
    )

    styles = getSampleStyleSheet()

    title = ParagraphStyle(
        "Title",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=24,
        textColor=colors.HexColor("#006747"),
        spaceAfter=10
    )

    heading = ParagraphStyle(
        "Heading",
        parent=styles["Heading2"],
        alignment=TA_CENTER,
        fontSize=16,
        textColor=colors.black
    )

    normal = ParagraphStyle(
        "NormalCenter",
        parent=styles["Normal"],
        alignment=TA_CENTER,
        fontSize=11
    )

    elements = []

    logo = "smartpay_logo/smartpay_logo.png"

    try:

        img = Image(
            logo,
            width=1.6*inch,
            height=1.6*inch
        )

        img.hAlign = "CENTER"

        elements.append(img)

    except:

        pass

    elements.append(Spacer(1,20))

    elements.append(
        Paragraph(
            "NATIONAL BANK OF PAKISTAN",
            title
        )
    )

    elements.append(
        Paragraph(
            "Digital Banking Group",
            heading
        )
    )

    elements.append(Spacer(1,20))

    elements.append(
        Paragraph(
            "SMARTPAY EXECUTIVE REPORT",
            title
        )
    )

    elements.append(Spacer(1,30))

    elements.append(
        Paragraph(
            f"<b>Prepared By:</b> Muhammad Saad Asif",
            normal
        )
    )

    elements.append(
        Paragraph(
            f"<b>Generated:</b> {datetime.now().strftime('%d %B %Y %I:%M %p')}",
            normal
        )
    )

    elements.append(Spacer(1,150))

    elements.append(
        Paragraph(
            "<font color='grey'>CONFIDENTIAL</font>",
            normal
        )
    )
    # =====================================
    # DASHBOARD SUMMARY
    # =====================================

    elements.append(PageBreak())

    elements.append(
        Paragraph(
            "Executive Dashboard Summary",
            title
        )
    )

    elements.append(Spacer(1,20))

    total = len(df)

    live = len(df[df["Status"].astype(str).str.upper()=="LIVE"])

    uat = len(df[df["Status"].astype(str).str.upper()=="UAT"])

    sit = len(df[df["Status"].astype(str).str.upper()=="SIT"])

    development = len(
        df[
            df["Status"]
            .astype(str)
            .str.upper()
            .str.contains("DEVELOP")
        ]
    )

    summary = [

        ["Metric","Value"],

        ["Total Projects", total],

        ["Live Projects", live],

        ["UAT Projects", uat],

        ["SIT Projects", sit],

        ["Development", development]

    ]

    table = Table(summary,colWidths=[250,120])

    table.setStyle(

        TableStyle([

            ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#006747")),

            ('TEXTCOLOR',(0,0),(-1,0),colors.white),

            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),

            ('BACKGROUND',(0,1),(-1,-1),colors.whitesmoke),

            ('GRID',(0,0),(-1,-1),0.5,colors.grey),

            ('BOTTOMPADDING',(0,0),(-1,0),10),

            ('TOPPADDING',(0,0),(-1,-1),8),

            ('ALIGN',(0,0),(-1,-1),'CENTER')

        ])

    )

    elements.append(table)

    elements.append(Spacer(1,25))
    # =====================================
    # STATUS ANALYTICS
    # =====================================

    elements.append(
        Paragraph(
            "Project Status Analytics",
            heading
        )
    )

    elements.append(Spacer(1,15))

    statuses = ["LIVE", "UAT", "SIT", "DEVELOPMENT"]

    values = [

        live,

        uat,

        sit,

        development

    ]

    drawing = Drawing(450,220)

    chart = VerticalBarChart()

    chart.x = 40

    chart.y = 30

    chart.height = 150

    chart.width = 320

    chart.data = [values]

    chart.categoryAxis.categoryNames = statuses

    chart.valueAxis.valueMin = 0

    chart.valueAxis.valueMax = max(values)+2

    chart.valueAxis.valueStep = 1

    chart.bars.fillColor = colors.HexColor("#006747")

    drawing.add(chart)

    title_lbl = Label()

    title_lbl.x = 160

    title_lbl.y = 200

    title_lbl.setText("Project Status Distribution")

    drawing.add(title_lbl)

    elements.append(drawing)

    elements.append(Spacer(1,25))
    
    # =====================================
    # TEAM PERFORMANCE
    # =====================================

    elements.append(PageBreak())

    elements.append(
        Paragraph(
            "Team Performance",
            title
        )
    )

    elements.append(Spacer(1,20))

    team = (
        df.groupby("Allocation")
        .size()
        .reset_index(name="Projects")
        .sort_values("Projects", ascending=False)
    )

    team_data = [["Team Member", "Projects"]]

    for _, row in team.iterrows():
        team_data.append([
            str(row["Allocation"]),
            str(row["Projects"])
        ])

    team_table = Table(
        team_data,
        colWidths=[260,120]
    )

    team_table.setStyle(
        TableStyle([

            ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#006747")),
            ('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),

            ('GRID',(0,0),(-1,-1),0.5,colors.grey),

            ('BACKGROUND',(0,1),(-1,-1),colors.whitesmoke),

            ('BOTTOMPADDING',(0,0),(-1,0),10),

            ('TOPPADDING',(0,1),(-1,-1),8),

            ('ALIGN',(1,1),(-1,-1),'CENTER')

        ])
    )

    elements.append(team_table)

    elements.append(Spacer(1,20))

    top = team.iloc[0]

    elements.append(

        Paragraph(

            f"<b>Top Performer :</b> {top['Allocation']} "
            f"({top['Projects']} Projects)",

            styles["Heading2"]

        )

    )

    elements.append(Spacer(1,25))

    # =====================================
    # PROJECT TIMELINE SUMMARY
    # =====================================

    elements.append(PageBreak())

    elements.append(
        Paragraph(
            "Project Timeline Summary",
            title
        )
    )

    elements.append(Spacer(1,20))

    timeline = [["Project", "Current Stage"]]

    for _, row in df.iterrows():

        stage = str(row["Status"]).upper()

        if stage == "LIVE":
            stage = "🟢 LIVE"

        elif stage == "UAT":
            stage = "🟡 UAT"

        elif stage == "SIT":
            stage = "🔵 SIT"

        elif "DEVELOP" in stage:
            stage = "🟠 DEVELOPMENT"

        timeline.append([
            str(row["Mandate"]),
            stage
        ])

    timeline_table = Table(
        timeline,
        colWidths=[270,130]
    )

    timeline_table.setStyle(
        TableStyle([

            ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#006747")),
            ('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),

            ('GRID',(0,0),(-1,-1),0.5,colors.grey),

            ('BACKGROUND',(0,1),(-1,-1),colors.beige),

            ('BOTTOMPADDING',(0,0),(-1,0),10),

            ('TOPPADDING',(0,1),(-1,-1),8)

        ])
    )

    elements.append(timeline_table)

    elements.append(Spacer(1,20))

    # =====================================
    # PROJECT DETAILS
    # =====================================

    elements.append(PageBreak())

    elements.append(
        Paragraph(
            "Detailed Project Information",
            title
        )
    )

    elements.append(Spacer(1,20))

    headers = [
        "Project",
        "Owner",
        "Status",
        "Category"
    ]

    table_data = [headers]

    for _, row in df.iterrows():

        table_data.append([

            str(row.get("Mandate", "")),

            str(row.get("Allocation", "")),

            str(row.get("Status", "")),

            str(row.get("Category", ""))

        ])

    project_table = Table(
        table_data,
        colWidths=[170,120,80,120],
        repeatRows=1
    )

    style = [

        ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#006747")),

        ('TEXTCOLOR',(0,0),(-1,0),colors.white),

        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),

        ('FONTSIZE',(0,0),(-1,0),10),

        ('GRID',(0,0),(-1,-1),0.3,colors.grey),

        ('BOTTOMPADDING',(0,0),(-1,0),10),

        ('TOPPADDING',(0,1),(-1,-1),7),

        ('BOTTOMPADDING',(0,1),(-1,-1),7),

        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),

    ]

    for i in range(1, len(table_data)):

        if i % 2 == 0:

            style.append(
                ('BACKGROUND',(0,i),(-1,i),colors.whitesmoke)
            )

        else:

            style.append(
                ('BACKGROUND',(0,i),(-1,i),colors.beige)
            )

    project_table.setStyle(TableStyle(style))

    elements.append(project_table)

    elements.append(Spacer(1,20))

    # =====================================
    # EXECUTIVE SUMMARY
    # =====================================

    elements.append(PageBreak())

    elements.append(
        Paragraph(
            "Executive Summary",
            title
        )
    )

    elements.append(Spacer(1,20))

    summary = f"""

    <b>Total Projects :</b> {total}<br/><br/>

    <b>Live Projects :</b> {live}<br/><br/>

    <b>UAT Projects :</b> {uat}<br/><br/>

    <b>SIT Projects :</b> {sit}<br/><br/>

    <b>Development :</b> {development}<br/><br/>

    This report was automatically generated from the
    SmartPay Project Dashboard developed for the
    Digital Banking Group, National Bank of Pakistan.

    """

    elements.append(
        Paragraph(
            summary,
            styles["BodyText"]
        )
    )

    elements.append(Spacer(1,30))

    elements.append(
        Paragraph(
            "<b>Prepared By</b><br/>Muhammad Saad Asif",
            heading
        )
    )

    elements.append(Spacer(1,20))

    elements.append(
        Paragraph(
            "<font color='#006747'><b>National Bank of Pakistan</b></font>",
            heading
        )
    )

    elements.append(
        Paragraph(
            "Digital Banking Group",
            normal
        )
    )

    elements.append(Spacer(1,80))

    elements.append(
        Paragraph(
            "<font color='grey'>*** End of Executive Report ***</font>",
            normal
        )
    )   

    doc.build(elements)

    buffer.seek(0)

    return buffer