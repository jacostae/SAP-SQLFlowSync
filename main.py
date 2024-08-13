from scripts import *

if __name__ == "__main__":
    
    Inicio_tiempo = start_timer()
    
    current_date = get_current_date_Colombia()

    yesterday_date_str, current_date_str, today_date_SAP, yesterday_date_SAP, start_month_str, end_month_str = calculate_dates_yesterday_today(current_date)

    df_Siclo_dia = current_data_Siclo(yesterday_date_str, current_date_str)

    df_F_dia = current_data_Facturacion(path_yesterday_today_data, yesterday_date_SAP, today_date_SAP)
    
    if yesterday_date_str == current_date_str:
        n = 1

    else:
        n = 2

    BD_Siclo, BD_SAP, posicion_ultimo_mes_D, posicion_ultimo_mes_F, worksheet_D, worksheet_F = data_spreadsheet(n)

    df_Siclo_final, df_Facturacion_final = join_data(df_Siclo_dia, BD_Siclo, df_F_dia, BD_SAP)
    
    df_cierre, registros_no_cierre, df_Siclo_final, df_Facturacion_final = match_Siclo_SAP(df_Siclo_final, df_Facturacion_final)
    
    df_concatenado = discarded_logs(registros_no_cierre)
    
    data, range_str, data_F, range_str_F, data_C, range_str_C = prepare_data(df_Siclo_final, posicion_ultimo_mes_D, df_Facturacion_final, posicion_ultimo_mes_F, df_cierre)
    
    update_data_spreadsheet(worksheet_D, worksheet_F, df_concatenado, current_date, data, range_str, data_F, range_str_F, data_C, range_str_C)

    tiempo = end_timer(Inicio_tiempo)
    
    save_time_log(current_date, tiempo, path_bitacora)