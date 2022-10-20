from turtle import back
from holidays import timedelta

import backend

import streamlit as st
import pandas as pd
import holidays
from datetime import date, datetime

################ Page Config ##################

st.set_page_config(page_title='PyCLOCK BOT', page_icon='🤖')

############### Global Variables ##############

y_today = date.today().year
m_today = date.today().month
d_today = date.today().day
target_daily_working_hours = backend.WEEKLY_WORKING_HOURS / backend.WEEKLY_WORKING_DAYS

visible_cols = ['Date', 'Day', 'Clock In',
                'Clock Out', 'Pause', 'Work Sum', 'Comment']

public_holidays = holidays.country_holidays(backend.COUNTRY, backend.STATE, y_today)

df_month = backend.load_month_report(y_today, m_today)
df_year = backend.load_year_report(y_today)
df_public_holidays = pd.DataFrame(public_holidays.items(), columns=['Date', 'Name'])

if 'Annual Leave' in df_year['Comment'].unique():
    annual_leave_remain = backend.ANNUAL_LEAVE_DAYS - (df_year['Comment'].value_counts()['Annual Leave']) 
else:
    annual_leave_remain = backend.ANNUAL_LEAVE_DAYS

############# Callback Functions ##############

def stamp_clock_in():
    """
    Write and save the clock in time of today
    """
    if df_month.loc[d_today-1, 'Comment'] != 'working day':
        status_box.warning('Not A Working Day')
        return

    df_month.loc[d_today-1, 'Clock In'] = datetime.now().strftime('%H:%M')

    backend.save_month_report(df_month, y_today, m_today)

    month_widget.dataframe(df_month[visible_cols], use_container_width=True)

    check_in_board.metric('Clock In', datetime.now().strftime('%H:%M'))

    status_box.success('Clock In Success :)')


def stamp_clock_out():
    """
    Write and save the clock out time of today
    """
    if df_month.loc[d_today-1, 'Comment'] != 'working day':
        status_box.warning('Not A Working Day')
        return

    t_clk_out = datetime.now().strftime('%H:%M')
    t_clk_in = df_month.loc[d_today-1, 'Clock In']
    t_pause = df_month.loc[d_today-1, 'Pause']

    df_month.loc[d_today-1, 'Clock Out'] = t_clk_out

    # total working hour = clock out - clout in - pause duration
    diff_1 = backend.calc_duration(t_clk_in, t_clk_out)
    t_working = backend.calc_duration(t_pause, diff_1)
    df_month.loc[d_today-1, 'Work Sum'] = t_working

    # update the worksum and overtime
    worksum = backend.calc_worksum(df_month, day=d_today)
    df_month.loc[df_month['Date'] == 'Summary', 'Work Sum'] = worksum

    overtime = backend.calc_overtime(df_month, day=d_today)
    month_board.markdown('**Overtime = {}**'.format(overtime))

    # save and update widgets
    backend.save_month_report(df_month, y_today, m_today)

    month_widget.dataframe(df_month[visible_cols], use_container_width=True)

    check_out_board.metric('Clock Out', datetime.now().strftime('%H:%M'))

    status_box.success('Clock Out Success :)')


def stamp_pause_begin():
    """
    Write and save the pause begin time
    """
    if df_month.loc[d_today-1, 'Comment'] != 'working day':
        status_box.warning('Not A Working Day')
        return

    t = datetime.now().strftime('%H:%M')

    df_month.loc[d_today-1, 'Pause Start'] = t

    backend.save_month_report(df_month, y_today, m_today)

    month_widget.dataframe(df_month[visible_cols], use_container_width=True)

    status_box.success('Break started :)')


def stamp_pause_end():
    """
    Write and save the pause end time
    """
    if df_month.loc[d_today-1, 'Comment'] != 'working day':
        status_box.warning('Not A Working Day')
        return

    t_end = datetime.now().strftime('%H:%M')
    t_begin = df_month.loc[d_today-1, 'Pause Start']

    # formatted pause duration in 'HH:MM'
    diff = backend.calc_duration(t_begin, t_end)
    prev_diff = df_month.loc[d_today-1, 'Pause']
    new_diff = backend.sum_time(diff, prev_diff)

    df_month.loc[d_today-1, 'Pause'] = new_diff
    df_month.loc[d_today-1, 'Pause Stop'] = t_end

    backend.save_month_report(df_month, y_today, m_today)

    month_widget.dataframe(df_month[visible_cols], use_container_width=True)

    pause_board.metric('Pause', value=df_month.loc[d_today-1, 'Pause'])

    status_box.success('Break ended :)')


def process_leave_apply(type_of_leave):
    """
    Update the tabular information for leave requests
    Input parameter type from 'Annual Leave', 'Sick Leave'
    """
    leave_days = (leave_stop_date - leave_start_date).days
    if leave_days < 0:
        status_box.error('Invalid Input: Check Your Date Selection Again')
        return

    valid_annual_leave_days = 0  # count the valid annual leave days (e.g. holidays, weekends are excluded)

    # get the related month_report
    dfs = []
    for i in range(leave_start_date.year, leave_stop_date.year+1):
        for j in range(leave_start_date.month, leave_stop_date.month+1):
            df = backend.load_month_report(i, j)
            dfs.append(df)

    for i in range(leave_days + 1):
        day_i = leave_start_date + timedelta(days=i)        # date object
        # find associate dataframe
        idx = (day_i.year - leave_start_date.year) + \
            (day_i.month - leave_start_date.month)
        df = dfs[idx]

        # update dateframe
        if df.loc[day_i.day-1, 'Comment'] == 'working day':
            df.loc[day_i.day-1, 'Comment'] = type_of_leave
            valid_annual_leave_days += 1 if type_of_leave == 'Annual Leave' else 0

    # save the workbook
    for df in dfs:
        y_df = int(df.loc[0, 'Date'][:4])
        m_df = int(df.loc[0, 'Date'][5:7])
        backend.save_month_report(df, y_df, m_df)

    df_month = backend.load_month_report(y_today, m_today)
    month_widget.dataframe(df_month[visible_cols], use_container_width=True)

    global annual_leave_remain 
    annual_leave_remain -= valid_annual_leave_days
    year_board.markdown('**Available annual leave days: {}**'.format(annual_leave_remain))


def process_leave_withdrawal(type_of_leave):
    """
    Update the tabular information for leave withdrawal requests
    Input parameter type from 'Annual Leave', 'Sick Leave'
    """
    leave_days = (leave_stop_date - leave_start_date).days
    if leave_days < 0:
        status_box.error('Invalid Input: Check Your Date Selection Again')
        return

    valid_annual_leave_days = 0  # count the valid annual leave days (e.g. holidays, weekends are excluded)
    
    # get the related month_report
    dfs = []
    for i in range(leave_start_date.year, leave_stop_date.year+1):
        for j in range(leave_start_date.month, leave_stop_date.month+1):
            df = backend.load_month_report(i, j)
            dfs.append(df)

    for i in range(leave_days + 1):
        day_i = leave_start_date + timedelta(days=i)        # date object
        # find associate dataframe
        idx = (day_i.year - leave_start_date.year) + \
            (day_i.month - leave_start_date.month)
        df = dfs[idx]

        # update dateframe
        if df.loc[day_i.day-1, 'Comment'] == 'Annual Leave':
            valid_annual_leave_days += 1
        df.loc[day_i.day-1, 'Comment'] = 'working day'

    # save the workbook
    for df in dfs:
        y_df = int(df.loc[0, 'Date'][:4])
        m_df = int(df.loc[0, 'Date'][5:7])
        backend.save_month_report(df, y_df, m_df)

    df_month = backend.load_month_report(y_today, m_today)
    month_widget.dataframe(df_month[visible_cols], use_container_width=True)

    global annual_leave_remain
    annual_leave_remain += valid_annual_leave_days
    year_board.markdown('**Available annual leave days: {}**'.format(annual_leave_remain))

###############################################
################## Main Page ##################
###############################################

st.title('PyCLOCK BOT 🤖')

############ Information Dashboard ############
st.header('Daily Info [{}]'.format(date.today().strftime('%Y-%m-%d')))

# Daily Info Board
col1, col2, col3 = st.columns(3)
with col1:
    check_in_board = st.metric(
        'Clock In', value=df_month.loc[d_today-1, 'Clock In'])
with col2:
    check_out_board = st.metric(
        'Clock Out', value=df_month.loc[d_today-1, 'Clock Out'])
with col3:
    pause_board = st.metric('Pause', value=df_month.loc[d_today-1, 'Pause'])

# Month & Year Info Board
col1, col2 = st.columns(2)
with col1:
    st.header('Overview {}'.format(date.today().strftime('%b-%Y')))
    ot = backend.calc_overtime(df_month, d_today)
    month_board = st.markdown('**Overtime = {}**'.format(ot))
with col2:
    st.header('Overview {}'.format(date.today().strftime('%Y')))
    year_board = st.markdown('**Available annual leave days: {}**'.format(annual_leave_remain))

# Holidays view expander
with st.expander('View Public Holidays This Year'):
    st.dataframe(df_public_holidays, use_container_width=True)

st.write('---')

############ Month Tabular View ############
with st.container():
    st.header('Monthly Overview')
    month_widget = st.dataframe(
        df_month[visible_cols], use_container_width=True)

##############################################
################## Side Bar ##################
##############################################

with st.sidebar:
    # Status Box
    with st.container():
        status_box = st.success('Welcome 📣📣')

    # LOGO
    from PIL import Image
    logo = Image.open('./img/打卡.png')
    st.image(logo)

    # Operation Area
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        btn_clock_in = st.button('CLOCK IN')
        if btn_clock_in:
            stamp_clock_in()
    with col2:
        btn_clock_out = st.button('CLOCK OUT')
        if btn_clock_out:
            stamp_clock_out()
    with col3:
        btn_start_pause = st.button('START BREAK')
        if btn_start_pause:
            stamp_pause_begin()
    with col4:
        btn_stop_pause = st.button('STOP BREAK', key='btn_stop')
        if btn_stop_pause:
            stamp_pause_end()

    st.write('---')

    # Leave apply and withdrawal
    st.subheader('Request Time Off')

    leave_start_date = st.date_input('From', date.today())
    leave_stop_date = st.date_input('To', date.today())

    col1, col2 = st.columns(2, gap='small')
    with col1:
        leave_check1 = st.checkbox('Annual Leave')
    with col2:
        leave_check2 = st.checkbox('Sick Leave')

    if leave_check1 and not leave_check2:
        leave_type = 'Annual Leave'
    elif leave_check2 and not leave_check1:
        leave_type = 'Sick Leave'

    col1, col2 = st.columns(2, gap='small')
    with col1:
        btn_apply = st.button('APPLY')
        if btn_apply:
            process_leave_apply(leave_type)
    with col2:
        btn_withdrawal = st.button('WITHDRAWAL')
        if btn_withdrawal:
            process_leave_withdrawal(leave_type)

    st.write('---')

    # Month report view
    with st.form('View'):
        st.subheader('View Month Records')
        y = st.slider('Year', 2022, 2025, value=y_today)
        m = st.slider('Month', 1, 12, value=m_today)
        btn_view = st.form_submit_button('View')

        if btn_view:
            df_selected = backend.load_month_report(y, m)
            month_widget.dataframe(
                df_selected[visible_cols], use_container_width=True)

