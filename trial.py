import pandas as pd
import smtplib
import matplotlib.pyplot as plt
import codecs
import os
import logging
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def calculate_kpis(df):
    df['LINES/ORDER'] = df['LINES'] / df['ORDERS']
    avg_ratio = '{:.2f} lines/order'.format(df['LINES/ORDER'].mean())
    max_ratio = '{:.2f} lines/order'.format(df['LINES/ORDER'].max())
    return avg_ratio, max_ratio

def generate_report(df, WEEK):
    busy_day = df.loc[df['LINES'].idxmax(), 'DAY']
    max_lines = '{:,} lines'.format(df['LINES'].max())
    total_lines = '{:,} lines'.format(df['LINES'].sum())
    return busy_day, max_lines, total_lines

def save_plot(df, WEEK):
    fig, ax = plt.subplots(figsize=(14, 7))
    df.plot.bar(x='DAY', y=['ORDERS', 'LINES'], color=['tab:blue', 'tab:red'], legend=True, ax=ax)
    plt.xlabel('DAY', fontsize=12)
    plt.title('Workload per day (Lines/day)', fontsize=12)
    plot_path = f'plot_{WEEK}.png'
    plt.savefig(plot_path, dpi=fig.dpi)
    return plot_path


def send_email(report_html, plot_path):
    # SMTP settings
    smtp_server = 'smtp.google.com'
    smtp_port = 465

    # Email settings
    from_mail = 'markorlando45@gmail.com'
    from_password = os.environ.get('EMAIL_PASSWORD')  # Use environment variable
    to_mail = 'yourmanager@gmail.com'

    msg = MIMEMultipart()
    msg['Subject'] = f'Workload Report for {WEEK}'
    msg['From'] = from_mail
    msg['To'] = to_mail
    msg.preamble = f'Workload Report for {WEEK}'

    with open(plot_path, 'rb') as fp:
        img = MIMEImage(fp.read())
        img.add_header('Content-Disposition', 'attachment', filename=plot_path)
        img.add_header('Content-ID', '<0>')
        msg.attach(img)

    # Populate template with data
    template = report_html.replace('WEEK', WEEK)
    template = template.replace('total_lines', total_lines)
    template = template.replace('busy_day', busy_day)
    template = template.replace('max_lines', max_lines)
    template = template.replace('avg_ratio', avg_ratio)
    template = template.replace('max_ratio', max_ratio)
    template = template.replace('IMG_HTML', 'cid:0')

    msg.attach(MIMEText(template, 'html', 'utf-8'))

    # Send the email
    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.ehlo()
        server.login(from_mail, from_password)
        server.sendmail(from_mail, to_mail, msg.as_string())
        server.quit()
        logging.info('Email sent successfully!')
    except Exception as e:
        logging.error('An error occurred while sending the email.')
        logging.error(str(e))

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)

    dict_days = {'MON': 'Monday', 'TUE': 'Tuesday', 'WED': 'Wednesday', 'THU': 'Thursday', 'FRI': 'Friday', 'SAT': 'Saturday', 'SUN': 'Sunday'}
    WEEK = 'WEEK-1'

    df_day = pd.read_csv('volumes_per_day.csv')
    df_plot = df_day[df_day['WEEK'] == WEEK].copy()

    avg_ratio, max_ratio = calculate_kpis(df_plot)
    busy_day, max_lines, total_lines = generate_report(df_plot, WEEK)
    plot_path = save_plot(df_plot, WEEK)
    
    with open("report_template.html", "r", encoding="utf-8") as f:
        report_html = f.read()

    send_email(report_html, plot_path)
