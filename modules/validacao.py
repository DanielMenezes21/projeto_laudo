from app.screen3_dadosp.dados_function import go_next
from app.screen4_territorio.solofunction import go_next1

VALIDACAO_ESPECIFICA = {
    'dados': lambda tela: go_next(tela),
    'territorio': lambda tela: go_next1(tela)
}
