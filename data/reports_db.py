# Dados simulados dos relatórios
# Futuramente será substituído por dados do Google Sheets

def get_report_data(client_id):
    """Retorna os dados do relatório para um cliente específico"""
    # Por enquanto, retorna os mesmos dados para todos os clientes
    # No futuro, buscará dados específicos por cliente do Google Sheets
    
    if client_id == 'enel':
        return {
            'table_data': [
                {
                    'indicador': 'Total demandado',
                    'acionados_2024': 24,
                    'acionados_2025': 153,
                    'total': 177,
                    'percentual': 100.0,
                    'is_main': True,
                    'is_total': True,
                    'level': 0
                },
                {
                    'indicador': 'Concluídos',
                    'acionados_2024': 22,
                    'acionados_2025': 147,
                    'total': 169,
                    'percentual': 95.5,
                    'is_main': True,
                    'is_total': False,
                    'level': 0
                },
                {
                    'indicador': 'Alvarás em andamento',
                    'acionados_2024': 2,
                    'acionados_2025': 6,
                    'total': 8,
                    'percentual': 4.5,
                    'is_main': True,
                    'is_total': False,
                    'level': 0
                },
                {
                    'indicador': 'Em análise Prefeitura',
                    'acionados_2024': None,
                    'acionados_2025': 1,
                    'total': 1,
                    'percentual': 0.5,
                    'is_main': False,
                    'is_total': False,
                    'level': 1
                },
                {
                    'indicador': 'Aguardando Pg. de Taxa Enel',
                    'acionados_2024': None,
                    'acionados_2025': None,
                    'total': None,
                    'percentual': None,
                    'is_main': False,
                    'is_total': False,
                    'level': 1
                },
                {
                    'indicador': 'Aguardando doc. Enel',
                    'acionados_2024': 1,
                    'acionados_2025': 3,
                    'total': 4,
                    'percentual': 2.2,
                    'is_main': False,
                    'is_total': False,
                    'level': 1
                },
                {
                    'indicador': 'Consulta de viabilidade',
                    'acionados_2024': None,
                    'acionados_2025': None,
                    'total': None,
                    'percentual': None,
                    'is_main': False,
                    'is_total': False,
                    'level': 1
                },
                {
                    'indicador': 'MR provi. doc. outros órgãos',
                    'acionados_2024': 1,
                    'acionados_2025': 2,
                    'total': 3,
                    'percentual': 1.8,
                    'is_main': False,
                    'is_total': False,
                    'level': 1
                },
                {
                    'indicador': 'Renov. de Al. prov. em andamento',
                    'acionados_2024': None,
                    'acionados_2025': None,
                    'total': None,
                    'percentual': None,
                    'is_main': False,
                    'is_total': False,
                    'level': 1
                }
            ],
            'chart_data': {
                'categories': ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio'],
                'values': [45, 52, 38, 61, 55]
            }
        }
    
    return None
