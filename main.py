import streamlit as st
import pandas as pd
import smartsheet



@st.cache_data(ttl=300)
def smartsheet_to_dataframe(sheet_id):
    smartsheet_client = smartsheet.Smartsheet(st.secrets['smartsheet']['access_token'])
    sheet             = smartsheet_client.Sheets.get_sheet(sheet_id)
    columns           = [col.title for col in sheet.columns]
    rows              = []
    for row in sheet.rows: rows.append([cell.value for cell in row.cells])
    return pd.DataFrame(rows, columns=columns)

@st.cache_data(ttl=300)
def fetch_smartsheet_dates():
    df           = smartsheet_to_dataframe(st.secrets['smartsheet']['sheets']['dates'])
    df           = df.sort_values(by=['Start','End']).reset_index(drop=True)
    df['Start']  = pd.to_datetime(df['Start']).dt.date
    df['End']    = pd.to_datetime(df['End']).dt.date
    return df



APP_NAME = 'Budget Buddy'

st.set_page_config(page_title=APP_NAME, page_icon='⚖️', layout='wide')

st.image(st.secrets['images']["rd_logo"], width=100)

st.title(APP_NAME)
st.info('Translation of the Booking Summary Report to ADR and utilization.')


file = st.file_uploader('Booking Summary Report | 01/01 - 12/31', type=['xlsx'])

if file:
    df                      = pd.read_excel(file, sheet_name='Sheet 1')
    df                      = df[['Unit_Code','First_Night','Last_Night','Nights','BookingRentTotal','ReservationTypeDescription']]
    df['First_Night']       = pd.to_datetime(df['First_Night'])
    df['Last_Night']        = pd.to_datetime(df['Last_Night'])
    df['ADR']               = df['BookingRentTotal'] / df['Nights']
    df                      = df[df['ReservationTypeDescription'] == 'Renter']
    df                      = df[['Unit_Code','First_Night','Last_Night','Nights','BookingRentTotal','ADR']]

    dates                   = fetch_smartsheet_dates()
    dates['Start']          = pd.to_datetime(dates['Start'])
    dates['End']            = pd.to_datetime(dates['End'])

    df['_key']              = 1
    dates['_key']           = 1
    merged                  = pd.merge(df, dates, on='_key').drop('_key', axis=1)

    merged['OverlapStart']  = merged[['First_Night', 'Start']].max(axis=1)
    merged['OverlapEnd']    = merged[['Last_Night', 'End']].min(axis=1)
    merged['OverlapDays']   = (merged['OverlapEnd'] - merged['OverlapStart']).dt.days + 1

    merged.loc[merged['OverlapDays'] < 0, 'OverlapDays'] = 0

    best    = merged.loc[merged.groupby(['First_Night', 'Last_Night'])['OverlapDays'].idxmax(),['First_Night', 'Last_Night', 'Comp Sets']]

    df      = df.merge(best, on=['First_Night', 'Last_Night'], how='left')
    df      = df.drop('_key', axis=1)

    agg     = df.groupby(['Unit_Code','Comp Sets']).agg(
        Nights      = ('Nights', 'sum'),
        ADR         = ('ADR', 'mean'),
        First_Night = ('First_Night', 'min'),
    ).reset_index()

    agg             = agg.sort_values(by=['Unit_Code','First_Night']).reset_index(drop=True)
    agg             = agg[['Unit_Code','Comp Sets','Nights','ADR']]

    liaisons        = smartsheet_to_dataframe(st.secrets['smartsheet']['sheets']['liaisons'])
    liaisons        = liaisons[['Unit_Code','OL']]

    final           = agg.merge(liaisons, on='Unit_Code', how='left')
    final['ADR']    = final['ADR'].round(0).astype(int)
    final['ADR']    = final['ADR'].apply(lambda x: f"${x:}")
    final.columns   = ['Unit Code', 'Season','Nights','ADR', 'OL']

    dates_agg = dates.groupby('Comp Sets', as_index=False).agg(
        Start = ('Start', 'min'),
        End   = ('End', 'max'),
    )

    dates_agg['Num of Days'] = (dates_agg['End'] - dates_agg['Start']).dt.days

    dates_agg = dates_agg[['Comp Sets', 'Num of Days']]
    dates_agg.columns = ['Season', 'Num of Days']
    final = final.merge(dates_agg, on='Season', how='left')
    final ['Booked'] = (final['Nights'] / final['Num of Days'] * 100).round(2).astype(str) + '%'

    final = final[['Unit Code', 'Season', 'ADR', 'Booked', 'OL']]
    st.dataframe(final, hide_index=True)