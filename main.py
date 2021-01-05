import sys
import pandas as pd


def main():
    try:
        df = pd.read_excel('data/sample01.xls', index_col=None, header=None)
        my_df = getFeesBreakdown(df)
        # Add new columns to the data frame
        my_df['invoice_no'] = getInvoiceDetails(df[7], 'Invoice # ')
        my_df['file_no'] = getInvoiceDetails(df[7], 'File Number: ')
        my_df['invoice_date'] = getInvoiceDetails(df[7], 'Invoice Date: ')
        # Write to csv
        my_df.to_csv('data/myfile.csv', index = False)
        print("Done")
    except Exception as error:
        print(error)

def getInvoiceDetails(df, keyword):
    """ To get invoice details from column P"""
    # Iterate each row in column P and search for the keyword
    for data in df:
        if keyword in str(data):
            # If keyword is found, only returns the value after the keyword
            return data.split(keyword)[-1]


def getFeesBreakdown(df):
    """ To get ifees breakdown"""
    # Use 'Fees' to identify to start row
    for row in range(df.shape[0]): 
        for col in range(df.shape[1]):
            if df.iat[row,col] == 'Fees':
                row_start = row
                break

    start_row_df = df.loc[row_start+1:]

    # remove the remaining rows if its date is NaN
    fees_df = start_row_df[:start_row_df[0].isnull().argmax()]

    # Replace header
    fees_df.columns = fees_df.iloc[0]
    fees_df = fees_df[1:]
    return fees_df.dropna(axis=1, how='all')


if __name__ == "__main__":
    sys.exit(main())
