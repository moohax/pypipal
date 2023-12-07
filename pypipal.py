import sys, os, argparse
import pandas as pd
import xlsxwriter

if sys.version_info.major != 3:
    print("[!] This script requires Python 3")
    sys.exit(1)

def main(args):

    #Read file
    data = pd.read_csv(args.infile, delimiter=args.sep, header=None, names=['Hash','Password'])

    #Build out master dataframe
    #basewords
    data['Baseword'] = data.Password.str.replace(r'[^a-zA-Z]*$', '')
    data['Baseword'] = data.Baseword.str.replace(r'^[^a-zA-Z]*', '')

    #password complexity
    data['loweralpha'] = data.Password.str.count(r'^[a-z]+$')
    data['upperalpha'] = data.Password.str.count(r'^[A-Z]+$')
    data['loweralphanum'] = data.Password.str.count(r'^[a-z0-9]+$')
    data['upperalphanum'] = data.Password.str.count(r'^[A-Z0-9]+$')
    data['mixedalpha'] = data.Password.str.count(r'^[a-zA-Z]+$')
    data['mixedalphanum'] = data.Password.str.count(r'^[a-zA-Z0-9]+$')
    data['mixedalphaspecial'] = data.Password.str.count(r'^[A-Za-z\\p{Punct}]+$')
    data['loweralphaspecialnum'] =  data.Password.str.count(r'^[a-z\\p{Punct}0-9]+$')
    data['upperalphaspecialnum'] = data.Password.str.count(r'^[A-Z\\p{Punct}0-9]+$')

    #password length
    data['Length'] = data['Password'].str.len()
    data = data.sort_values('Length', ascending=False).reset_index(drop=True)

    #Describe the data
    d_describe = data.Password.describe(include='all').to_string()
    print('Here\'s your data!')
    print('Description:\n{}\r\n'.format(d_describe))

    #Build dataframes for excel
    #Cracked uncracked, total
    print('[+] Creating first dataframe...')
    df1 = pd.DataFrame([[data.Hash.count(), data.Password.count(), data.Hash.count() - data.Password.count()]],  columns=('Total', 'Cracked', 'Uncracked')).T
    
    #Top 10 passwords
    print('[+] Creating second dataframe...')
    df2 = pd.DataFrame([data.Password.value_counts()[:10]]).T

    #Password length distribution
    print('[+] Creating third dataframe...')
    df3 = pd.DataFrame([data.Length.value_counts()]).T

    #Top 10 base words
    print('[+] Creating fourth dataframe...')
    df4 = pd.DataFrame([data.Baseword.value_counts()[:10]]).T

    #Character analysis
    print('[+] Creating fifth dataframe...')
    df5 = pd.DataFrame([[data.loweralpha.sum(), data.upperalpha.sum(), data.loweralphanum.sum(), data.upperalphanum.sum(), data.mixedalpha.sum(), data.mixedalphanum.sum(), data.mixedalphaspecial.sum(), data.loweralphaspecialnum.sum(), data.upperalphaspecialnum.sum()]],
    columns=('loweralpha', 'upperalpha', 'loweralphanum', 'upperalphanum', 'mixedalpha', 'mixedalphanum', 'mixedalphaspecial', 'loweralphaspecialnum', 'upperalphaspecialnum')).T
    print('[+] Dataframes created!')

    # Create workbook.
    print('[+] Writing the Excel charts...')
    writer = pd.ExcelWriter(args.outfile, engine='xlsxwriter')
    df1.to_excel(writer, sheet_name='Password_Analysis')
    df2.to_excel(writer, sheet_name='Password_Analysis', startrow=6)
    df3.to_excel(writer, sheet_name='Password_Analysis', startrow=22)
    df4.to_excel(writer, sheet_name='Password_Analysis', startrow=48)
    df5.to_excel(writer, sheet_name='Password_Analysis', startrow=68)

    workbook  = writer.book
    worksheet = writer.sheets['Password_Analysis']

    #Steal Trevors code for charts
    # Percentage Cracked
    print('[+] Writing first chart...') 
    chart1 = workbook.add_chart({'type': 'pie'})
    chart1.set_chartarea({
        'border': {'color': '#d9d9d9'}
        })

    # Configure graph details, including range of data to project graph with.
    chart1.add_series({
        'name': '# of passwords',
        'categories': '=Password_Analysis!$A$3:$A$4',
        'values':     '=Password_Analysis!$B$3:$B$4',
        'data_labels': {'value': True, 'separator': '\n', 'percentage': True},
        'points': [
            {'fill': {'color': '#5b9bd5'}},
            {'fill': {'color': '#f79646'}}
            ]
        })

    # Add a chart title.
    chart1.set_title ({
        'name': 'Percentage Cracked',
        'name_font': {
            'color': '#595959'
            }
        })

    # Set an excel chart style.
    chart1.set_style(24)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D1', chart1, {'x_offset': 25, 'y_offset': 10})

    #Top 10 Passwords
    print('[+] Writing second chart...')
    chart2 = workbook.add_chart({'type': 'bar'})
    chart2.set_chartarea({
        'border': {'color': '#d9d9d9'}
        })

    chart2.set_x_axis({
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#d9d9d9'}
        }
    })

    # Configure graph details, including range of data to project graph with.
    chart2.add_series({
        'name':       '# of passwords',
        'categories': '=Password_Analysis!$A$8:$A$17',
        'values':     '=Password_Analysis!$B$8:$B$17',
        'data_labels': {'value': True},
        'font': {'name': 'Calibri (Body)'},
        'fill':   {'color': '#5b9bd5'}
    })

    # Add a chart title.
    chart2.set_title ({
        'name': 'Top 10 Passwords',
        'name_font': {
            'color': '#595959'
        }
    })

    # Set an excel chart style.
    chart2.set_style(18)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D17', chart2, {'x_offset': 25, 'y_offset': 10})

    #Password length
    print('[+] Writing third chart...')
    chart3 = workbook.add_chart({'type': 'column'})
    chart3.set_chartarea({
        'border': {'color': '#d9d9d9'}
    })

    chart3.set_y_axis({
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#d9d9d9'}
        }
    })

    # Configure graph details, including range of data to project graph with.
    chart3.add_series({
        'name':       '# of passwords',
        'categories':     '=Password_Analysis!$A$24:$A$35',
        'values':     '=Password_Analysis!$B$24:$B$35',
        'data_labels': {'value': True},
        'font': {'name': 'Calibri (Body)'},
        'fill':   {'color': '#5b9bd5'}
    })

    # Add a chart title.
    chart3.set_title ({
        'name': 'Password Length',
        'name_font': {
            'color': '#595959'
        }
    })


    # Set an excel chart style.
    chart3.set_style(18)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D35', chart3, {'x_offset': 25, 'y_offset': 10})


    #Basewords
    print('[+] Writing fourth chart...')
    chart4 = workbook.add_chart({'type': 'bar'})
    chart4.set_chartarea({
        'border': {'color': '#d9d9d9'}
    })

    chart4.set_x_axis({
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#d9d9d9'}
        }
    })

    # Configure graph details, including range of data to project graph with.
    chart4.add_series({
        'name':       '# of base-words',
        'categories': '=Password_Analysis!$A$50:$A$59',
        'values':     '=Password_Analysis!$B$50:$B$59',
        'data_labels': {'value': True},
        'font': {'name': 'Calibri (Body)'},
        'fill':   {'color': '#5b9bd5'}
    })

    # Add a chart title.
    chart4.set_title ({
        'name': 'Top 10 Base Words',
        'name_font': {
            'color': '#595959'
        }
    })

    # Set an excel chart style.
    chart4.set_style(18)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D50', chart4, {'x_offset': 25, 'y_offset': 10})

    #Password complexity
    print('[+] Writing fifth chart...')
    chart5 = workbook.add_chart({'type': 'column'})
    chart5.set_chartarea({
        'border': {'color': '#d9d9d9'}
    })

    chart5.set_y_axis({
        'major_gridlines': {
            'visible': True,
            'line': {'color': '#d9d9d9'}
        }
    })

    # Configure graph details, including range of data to project graph with.
    chart5.add_series({
        'name':       '# of passwords',
        'categories': '=Password_Analysis!$A$70:$A$79',
        'values':     '=Password_Analysis!$B$70:$B$79',
        'data_labels': {'value': True},
        'font': {'name': 'Calibri (Body)'},
        'fill':   {'color': '#5b9bd5'}
    })

    # Add a chart title.
    chart5.set_title ({
        'name': 'Password Complexity',
        'name_font': {
            'color': '#595959'
        }
    })

    # Set an excel chart style.
    chart5.set_style(18)

    chart5.set_x_axis({
        'num_font': {'rotation': -45}
    })

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart('D70', chart5, {'x_offset': 25, 'y_offset': 10})
    print('[+] Excel written, saving...')

    writer.save()
    print('[+] Here you go, {}'.format(args.outfile))

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Pipal password processesing', formatter_class=argparse.RawDescriptionHelpFormatter, add_help=True)
    parser.add_argument('-f','--infile', help='Password file, H2P output, both cracked and uncracked in the same file \n', required=True)
    parser.add_argument('-s', '--sep', help='Separator for data column default is : (Should not need to change for H2P)', default=':')
    parser.add_argument('-o', '--outfile', help='Specify an outfile.xlsx', required=True)
    args = parser.parse_args()
    main(args)
