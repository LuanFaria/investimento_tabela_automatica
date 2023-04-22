from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.options import Options


class planilha_automatizada:

    def inicializar(self):

        self.configurar_web_driver()
        self.definir_url()
        self.capturar_valores_btra()
        self.capturar_valores_rbff()
        self.capturar_valores_bcff()
        self.capturar_valores_hsml()
        self.capturar_valores_visc()
        self.capturar_valores_bbpo()
        self.capturar_valores_vino()
        self.capturar_valores_vghf()
        self.capturar_valores_mxrf()
        self.capturar_valores_kncr()
        self.capturar_valores_bcri()
        self.capturar_valores_hglg()
        self.capturar_valores_ggrc()
        self.buscar_tabela_excel()

        print('\033[32m', '                FINALIZADO')
        print('\033[m', '')

    def configurar_web_driver(self):

        self.fire_op = Options()
        self.fire_op.headless = True
        self.firefox = webdriver.Firefox(options=self.fire_op)
        print('')
        print('\033[33m', 'Iniciando captura automatica de Dy e V/Vp')
        print('')

        # caso queira ver o processo apagar as 3 ultimas linhas e usar a de baixo
        # self.firefox = webdriver.Firefox()

    def definir_url(self):

        self.btra_url = 'https://statusinvest.com.br/fundos-imobiliarios/btra11'
        self.rbff_url = 'https://statusinvest.com.br/fundos-imobiliarios/rbff11'
        self.bcff_url = 'https://statusinvest.com.br/fundos-imobiliarios/bcff11'
        self.hsml_url = 'https://statusinvest.com.br/fundos-imobiliarios/hsml11'
        self.visc_url = 'https://statusinvest.com.br/fundos-imobiliarios/visc11'
        self.bbpo_url = 'https://statusinvest.com.br/fundos-imobiliarios/bbpo11'
        self.vino_url = 'https://statusinvest.com.br/fundos-imobiliarios/vino11'
        self.vghf_url = 'https://statusinvest.com.br/fundos-imobiliarios/vghf11'
        self.mxrf_url = 'https://statusinvest.com.br/fundos-imobiliarios/mxrf11'
        self.kncr_url = 'https://statusinvest.com.br/fundos-imobiliarios/kncr11'
        self.bcri_url = 'https://statusinvest.com.br/fundos-imobiliarios/bcri11'
        self.hglg_url = 'https://statusinvest.com.br/fundos-imobiliarios/hglg11'
        self.ggrc_url = 'https://statusinvest.com.br/fundos-imobiliarios/ggrc11'

    def capturar_valores_btra(self):

        self.firefox.get(self.btra_url)
        self.dy_btra = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_btra = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_btra = self.dy_btra.replace(',', '.')
        self.v_vp_btra = self.v_vp_btra.replace(',', '.')
        # print(self.dy_btra, self.v_vp_btra)
        self.firefox.quit
        print('\033[32m', 'Valores de BTRA - OK')

    def capturar_valores_rbff(self):
        self.firefox.get(self.rbff_url)
        self.dy_rbff = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_rbff = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_rbff = self.dy_rbff.replace(',', '.')
        self.v_vp_rbff = self.v_vp_rbff.replace(',', '.')
        self.firefox.quit
        print('\033[32m', 'Valores de RBFF - OK')

    def capturar_valores_bcff(self):
        self.firefox.get(self.bcff_url)
        self.dy_bcff = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_bcff = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_bcff = self.dy_bcff.replace(',', '.')
        self.v_vp_bcff = self.v_vp_bcff.replace(',', '.')
        self.firefox.quit
        print('\033[32m', 'Valores de BCFF - OK')

    def capturar_valores_hsml(self):
        self.firefox.get(self.hsml_url)
        self.dy_hsml = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_hsml = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_hsml = self.dy_hsml.replace(',', '.')
        self.v_vp_hsml = self.v_vp_hsml.replace(',', '.')
        self.firefox.quit
        print('\033[32m', 'Valores de HSML - OK')

    def capturar_valores_visc(self):
        self.firefox.get(self.visc_url)
        self.dy_visc = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_visc = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_visc = self.dy_visc.replace(',', '.')
        self.v_vp_visc = self.v_vp_visc.replace(',', '.')
        self.firefox.quit
        print('\033[32m', 'Valores de VISC - OK')

    def capturar_valores_bbpo(self):
        self.firefox.get(self.bbpo_url)
        self.dy_bbpo = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_bbpo = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_bbpo = self.dy_bbpo.replace(',', '.')
        self.v_vp_bbpo = self.v_vp_bbpo.replace(',', '.')
        self.firefox.quit
        print('\033[32m', 'Valores de BBPO - OK')

    def capturar_valores_vino(self):
        self.firefox.get(self.vino_url)
        self.dy_vino = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_vino = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_vino = self.dy_vino.replace(',', '.')
        self.v_vp_vino = self.v_vp_vino.replace(',', '.')
        self.firefox.quit
        print('\033[32m', 'Valores de VINO - OK')

    def capturar_valores_vghf(self):
        self.firefox.get(self.vghf_url)
        self.dy_vghf = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_vghf = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_vghf = self.dy_vghf.replace(',', '.')
        self.v_vp_vghf = self.v_vp_vghf.replace(',', '.')
        self.firefox.quit
        print('\033[32m', 'Valores de VGHF - OK')

    def capturar_valores_mxrf(self):
        self.firefox.get(self.mxrf_url)
        self.dy_mxrf = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_mxrf = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_mxrf = self.dy_mxrf.replace(',', '.')
        self.v_vp_mxrf = self.v_vp_mxrf.replace(',', '.')
        self.firefox.quit
        print('\033[32m', 'Valores de MXRF - OK')

    def capturar_valores_kncr(self):
        self.firefox.get(self.kncr_url)
        self.dy_kncr = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_kncr = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_kncr = self.dy_kncr.replace(',', '.')
        self.v_vp_kncr = self.v_vp_kncr.replace(',', '.')
        self.firefox.quit
        print('\033[32m', 'Valores de KNCR - OK')

    def capturar_valores_bcri(self):
        self.firefox.get(self.bcri_url)
        self.dy_bcri = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_bcri = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_bcri = self.dy_bcri.replace(',', '.')
        self.v_vp_bcri = self.v_vp_bcri.replace(',', '.')
        self.firefox.quit
        print('\033[32m', 'Valores de BCRI - OK')

    def capturar_valores_hglg(self):
        self.firefox.get(self.hglg_url)
        self.dy_hglg = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_hglg = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_hglg = self.dy_hglg.replace(',', '.')
        self.v_vp_hglg = self.v_vp_hglg.replace(',', '.')
        self.firefox.quit
        print('\033[32m', 'Valores de HGLG - OK')

    def capturar_valores_ggrc(self):
        self.firefox.get(self.ggrc_url)
        self.dy_ggrc = self.firefox.find_element(
            By.XPATH, '//*[@id="main-2"]/div[2]/div[1]/div[4]/div/div[1]/strong').text
        self.v_vp_ggrc = self.firefox.find_element(
            By.XPATH, '/html/body/main/div[2]/div[5]/div/div[2]/div/div[1]/strong').text
        self.dy_ggrc = self.dy_ggrc.replace(',', '.')
        self.v_vp_ggrc = self.v_vp_ggrc.replace(',', '.')
        self.firefox.quit
        print('\033[32m', 'Valores de GGRC - OK')
        print('')
        print('')
        print('\033[33m', '--------------- Salvando tabela Excel ---------------')
        print('')
        print('')

    def buscar_tabela_excel(self):

        import openpyxl

        # self.nome_do_arquivo = 'Investimento.xlsx'
        # self.dados = pd.read_excel(self.nome_do_arquivo, sheet_name='planilha_automatica')

        self.nome_do_arquivo = openpyxl.load_workbook('Investimento.xlsx')
        self.dados = self.nome_do_arquivo['planilha_automatica']

        for rows in self.dados.iter_rows(min_row=2, max_row=14):
            # print(f'{rows[0].value} {rows[1].value} {rows[2].value}')

            if rows[0].value == 'BTRA':  # Nome do Fundo
                rows[1].value = self.v_vp_btra  # Valor sobre valor patrimonial
                rows[2].value = self.dy_btra  # Porcentagem do dividendo anual

            if rows[0].value == 'RBFF':  # Nome do Fundo
                rows[1].value = self.v_vp_rbff  # Valor sobre valor patrimonial
                rows[2].value = self.dy_rbff  # Porcentagem do dividendo anual

            if rows[0].value == 'BCFF':  # Nome do Fundo
                rows[1].value = self.v_vp_bcff  # Valor sobre valor patrimonial
                rows[2].value = self.dy_bcff  # Porcentagem do dividendo anual

            if rows[0].value == 'HSML':  # Nome do Fundo
                rows[1].value = self.v_vp_hsml  # Valor sobre valor patrimonial
                rows[2].value = self.dy_hsml  # Porcentagem do dividendo anual

            if rows[0].value == 'VISC':  # Nome do Fundo
                rows[1].value = self.v_vp_visc  # Valor sobre valor patrimonial
                rows[2].value = self.dy_visc  # Porcentagem do dividendo anual

            if rows[0].value == 'BBPO':  # Nome do Fundo
                rows[1].value = self.v_vp_bbpo  # Valor sobre valor patrimonial
                rows[2].value = self.dy_bbpo  # Porcentagem do dividendo anual

            if rows[0].value == 'VINO':  # Nome do Fundo
                rows[1].value = self.v_vp_vino  # Valor sobre valor patrimonial
                rows[2].value = self.dy_vino  # Porcentagem do dividendo anual

            if rows[0].value == 'VGHF':  # Nome do Fundo
                rows[1].value = self.v_vp_vghf  # Valor sobre valor patrimonial
                rows[2].value = self.dy_vghf  # Porcentagem do dividendo anual

            if rows[0].value == 'MXRF':  # Nome do Fundo
                rows[1].value = self.v_vp_mxrf  # Valor sobre valor patrimonial
                rows[2].value = self.dy_mxrf  # Porcentagem do dividendo anual

            if rows[0].value == 'KNCR':  # Nome do Fundo
                rows[1].value = self.v_vp_kncr  # Valor sobre valor patrimonial
                rows[2].value = self.dy_kncr  # Porcentagem do dividendo anual

            if rows[0].value == 'BCRI':  # Nome do Fundo
                rows[1].value = self.v_vp_bcri  # Valor sobre valor patrimonial
                rows[2].value = self.dy_bcri  # Porcentagem do dividendo anual

            if rows[0].value == 'HGLG':  # Nome do Fundo
                rows[1].value = self.v_vp_hglg  # Valor sobre valor patrimonial
                rows[2].value = self.dy_hglg  # Porcentagem do dividendo anual

            if rows[0].value == 'GGRC':  # Nome do Fundo
                rows[1].value = self.v_vp_ggrc  # Valor sobre valor patrimonial
                rows[2].value = self.dy_ggrc  # Porcentagem do dividendo anual

            # print(f'{rows[0].value} {rows[1].value} {rows[2].value}')

        self.nome_do_arquivo.save('Investimento.xlsx')


start = planilha_automatizada()
start.inicializar()
