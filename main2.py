import re
import time
from collections import OrderedDict
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import load_workbook, Workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, ElementClickInterceptedException

# ================== ENTRADA ==================
documento = input("digite o nome do documento: ").strip()

# Planilha de alunos (entrada)
excel = load_workbook('ALUNOS.xlsx')
listaDeAlunos = excel['Planilha1']

# ================== PLANILHA DE SAÍDA ==================
OUT_PATH = 'RESULTADO.xlsx'
COL_FIXA_F = "Histórico Escolar do Ensino Fundamental"
COL_FIXA_G = "Declaração de Conclusão do Ensino Fundamental"
COL_FIXA_H = "Declaração de Vacinação Atualizada"
COL_FIXA_I = "Histórico Escolar Parcial para Transferência Externa - Ensino Médio Técnico"

try:
    wb_out = load_workbook(OUT_PATH)
    ws_out = wb_out.active
except Exception:
    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.append([
        'Nome', 'ID', 'CPF', 'Total de Doc', 'Status Doc da Pesquisa',
        COL_FIXA_F, COL_FIXA_G, COL_FIXA_H, COL_FIXA_I, 'Documento 1'
    ])

# ================== SELENIUM ==================
driver = webdriver.Chrome()
driver.set_window_size(1440, 900)  # evita layout mobile/accordion
driver.get('https://www.intranet.sp.senac.br/home')

def colocar_assim_aparecer(tipo, nome, conteudo, timeout=15):
    try:
        elemento = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((tipo, nome))
        )
        try:
            elemento.clear()
        except Exception:
            pass
        elemento.send_keys(str(conteudo))
        print("✅ Conteúdo inserido com sucesso.")
        return elemento
    except Exception as e:
        print(f"Erro ao inserir conteúdo '{conteudo}': {e}")
        raise

def clicar_assim_aparecer(tipo, nome, timeout=15):
    try:
        elemento = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((tipo, nome))
        )
        try:
            elemento.click()
        except Exception:
            driver.execute_script("arguments[0].click();", elemento)
        print("✅ Clicou com sucesso:", nome)
        return elemento
    except Exception as e:
        print(f"Erro ao clicar no elemento {nome}: {e}")
        raise

# ================== NAVEGAÇÃO INICIAL (como no seu script) ==================
clicar_assim_aparecer(By.XPATH, '/html/body/header/section/nav/ul/li[1]/a')
clicar_assim_aparecer(By.XPATH, '//*[@id="listagemPessoas"]/div/div[5]/div/div/div[2]/div/div[1]/div[1]/a')

handles_antes = driver.window_handles.copy()
clicar_assim_aparecer(By.XPATH, '//*[@id="divDetalhesSistema"]/div[2]/div[1]/a')
WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(len(handles_antes) + 1))
driver.switch_to.window(driver.window_handles[-1])

clicar_assim_aparecer(By.XPATH, '//*[@id="app"]/div/div/div[1]/div/div[2]/div[1]/button[2]/div')
clicar_assim_aparecer(By.XPATH, '//*[@id="item_10282CD6E9BF8D4E23F68971406D4AA4"]/div/li/a/span')

# ================== HELPERS ==================
def voltar_para_pesquisa(timeout=20):
    """Volta ao formulário de busca e espera o campo ficar clicável."""
    try:
        clicar_assim_aparecer(By.XPATH, '//*[@id="app"]/div/div/div[1]/div/div[2]/div[1]/button[2]/div', timeout=timeout)
        clicar_assim_aparecer(By.XPATH, '//*[@id="item_10282CD6E9BF8D4E23F68971406D4AA4"]/div/li/a/span', timeout=timeout)
    except Exception:
        pass
    WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[data-v-bc1d237e]')))

def estou_na_pagina_documentos(timeout=2):
    try:
        WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.XPATH, '//h2[normalize-space()="Documentos do Prontuário Educacional"]'))
        )
        return True
    except Exception:
        return False

def clicar_visualizar_primeira_linha(timeout=20, tentativas=3):
    """Clica no botão Visualizar da primeira linha da grade de busca."""
    # espera pelo menos 1 linha na grid
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '#sn-table-desk tbody tr'))
    )

    seletores = [
        # 1) botão de ação padrão
        (By.XPATH, '//*[@id="sn-table-desk"]/tbody/tr[1]/td[4]//button'),
        # 2) por data-cy
        (By.CSS_SELECTOR, '#sn-table-desk tbody tr:nth-child(1) td:nth-child(4) button[data-cy="read"]'),
        # 3) por ícone alt
        (By.XPATH, '//*[@id="sn-table-desk"]/tbody/tr[1]/td[4]//button[.//img[@alt="Visualizar"]]'),
        # 4) por ícone pelo nome do arquivo do SVG
        (By.XPATH, '//*[@id="sn-table-desk"]/tbody/tr[1]/td[4]//button[.//img[contains(@src,"search-sm-3329cf7e.svg")]]'),
        # 5) seletor amplo (qualquer linha)
        (By.XPATH, '//*[@id="sn-table-desk"]/tbody/tr/td[4]/button'),
    ]

    for _ in range(tentativas):
        for by, sel in seletores:
            try:
                btn = WebDriverWait(driver, 4).until(EC.element_to_be_clickable((by, sel)))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                try:
                    btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", btn)

                # confirma que entrou na página de documentos
                if estou_na_pagina_documentos(timeout=6):
                    print("✅ Entrou na página de documentos.")
                    return True
            except TimeoutException:
                continue
            except (ElementClickInterceptedException, StaleElementReferenceException):
                time.sleep(0.5)
                continue

    raise TimeoutException("Não foi possível clicar em 'Visualizar'.")

def abrir_dropdown_todos_documentos(timeout=12):
    """Abre Mostrar Mais → Todos APENAS na página de documentos."""
    if not estou_na_pagina_documentos():
        return
    try:
        btn = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((
            By.XPATH,
            '//div[contains(@class,"dropdown")][.//span[normalize-space()="Mostrar Mais"]]//button'
        )))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)

        menu = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((
            By.XPATH,
            '//div[contains(@class,"dropdown")][.//span[normalize-space()="Mostrar Mais"]]//ul[contains(@class,"options")]'
        )))
        opc_todos = menu.find_element(By.XPATH, './/a[normalize-space()="Todos"]')
        try:
            ActionChains(driver).move_to_element(opc_todos).click().perform()
        except Exception:
            driver.execute_script("arguments[0].click();", opc_todos)
    except Exception:
        pass

def garantir_pagina_prontuario(timeout=15):
    if estou_na_pagina_documentos(timeout=2):
        return
    try:
        btn = driver.find_element(By.CSS_SELECTOR, '.collapse-profile-button')
        driver.execute_script("arguments[0].click();", btn)
    except Exception:
        pass
    WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((By.XPATH, '//main//div[contains(@class,"sn-spacing")]'))
    )

def _get_text_by_label(driver, label, timeout=10):
    xp_desk = f'//div[contains(@class,"sn-profile-details")]//h3[normalize-space()="{label}"]/following-sibling::h4[1]'
    try:
        el = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, xp_desk)))
        return el.text.strip()
    except Exception:
        pass
    xp_mob = f'//div[contains(@class,"mobile-profile-details")]//h3[contains(@class,"mobile-title-label")][normalize-space()="{label}"]/following-sibling::h4[contains(@class,"mobile-value-label")][1]'
    el = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, xp_mob)))
    return el.text.strip()

def _parse_docs_desktop(driver):
    docs = []
    rows = driver.find_elements(By.CSS_SELECTOR, '#sn-table-desk tbody tr')
    for r in rows:
        try:
            tds = r.find_elements(By.CSS_SELECTOR, 'td')
            tipo   = tds[0].find_element(By.CSS_SELECTOR, 'span').text.strip()
            origem = tds[1].find_element(By.CSS_SELECTOR, 'span').text.strip()
            data   = tds[2].find_element(By.CSS_SELECTOR, 'span').text.strip()
            docs.append({"tipo": tipo, "origem": origem, "data": data})
        except Exception:
            continue
    return docs

def _parse_docs_mobile(driver):
    docs = []
    titles = driver.find_elements(By.CSS_SELECTOR, '#sn-table-mobile .accordion-button')
    for t in titles:
        tipo = t.text.strip()
        try:
            ancestor = t.find_element(By.XPATH, './ancestor::tr[1]/following-sibling::div[1]')
            origem_el = ancestor.find_element(By.XPATH, './/td[contains(@class,"tbl-accordion-body")][@data-title="Origem :"]/span')
            data_el   = ancestor.find_element(By.XPATH, './/td[contains(@class,"tbl-accordion-body")][@data-title="Data de Análise :"]/span')
            docs.append({"tipo": tipo, "origem": origem_el.text.strip(), "data": data_el.text.strip()})
        except Exception:
            continue
    return docs

def scrape_prontuario(driver, timeout=15):
    garantir_pagina_prontuario(timeout=timeout)
    # título da página deve indicar documentos
    WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((By.XPATH, '//h2[normalize-space()="Documentos do Prontuário Educacional"]'))
    )

    nome   = _get_text_by_label(driver, 'Nome', timeout)
    cpf    = _get_text_by_label(driver, 'CPF', timeout)
    emplid = _get_text_by_label(driver, 'EMPLID', timeout)
    total  = _get_text_by_label(driver, 'Total de Documentos', timeout)

    documentos = []
    try:
        WebDriverWait(driver, 4).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#sn-table-desk')))
        documentos = _parse_docs_desktop(driver)
    except Exception:
        try:
            WebDriverWait(driver, 4).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#sn-table-mobile')))
            documentos = _parse_docs_mobile(driver)
        except Exception:
            documentos = []

    return {"nome": nome, "id": emplid, "cpf": cpf, "total_documentos": total, "documentos": documentos}

# ================== NORMALIZAÇÃO / AGRUPAMENTO ==================
REV_REGEX = re.compile(r'\s*\|\s*Revis[aã]o:\s*\d+\s*$', re.IGNORECASE)

def normaliza_titulo(t):
    if not t:
        return ''
    t = REV_REGEX.sub('', t).strip()
    return t

def so_digitos(s):
    m = re.search(r'\d+', s or '')
    return int(m.group()) if m else 0

def agrega_docs(docs):
    """Lista agregada com contagem (X) para repetidos."""
    ordem = OrderedDict()
    for d in docs:
        nome = normaliza_titulo(d.get('tipo', ''))
        if not nome:
            continue
        ordem[nome] = ordem.get(nome, 0) + 1
    saida = []
    for nome, qtd in ordem.items():
        saida.append(f"{nome} ({qtd})" if qtd > 1 else nome)
    return saida, ordem  # também devolve o dicionário p/ checagem rápida

def status_pesquisa(doc_pesquisado, lista_docs_normalizados):
    alvo = normaliza_titulo(doc_pesquisado).casefold()
    return "OK" if any(normaliza_titulo(x).casefold() == alvo for x in lista_docs_normalizados) else "Não Encontrado"

def garantir_cabecalho_docs(n_docs_extra):
    """Garante cabeçalhos 'Documento N' a partir da coluna J (índice 10)."""
    headers = [cell.value for cell in ws_out[1]]
    base_fixas = 9  # A..E + F..I = 9 colunas antes dos extras
    num_atual = max(0, len(headers) - (base_fixas))
    if n_docs_extra > num_atual:
        for i in range(num_atual + 1, n_docs_extra + 1):
            ws_out.cell(row=1, column=base_fixas + i, value=f"Documento {i}")

# ================== LOOP PRINCIPAL ==================
for row in listaDeAlunos.iter_rows(min_row=2, max_row=listaDeAlunos.max_row, min_col=1, max_col=listaDeAlunos.max_column):
    current_id = row[1].value
    if not current_id:
        print("Linha com ID vazio. Pulando.")
        continue

    try:
        # BUSCA PELO EMPLID
        colocar_assim_aparecer(By.CSS_SELECTOR, 'input[data-v-bc1d237e]', current_id)
        clicar_assim_aparecer(By.XPATH, '//*[@id="app"]/div/div/div[2]/main/div/div/div[2]/div/div[1]/div/button')

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#sn-table-desk tbody tr'))
        )

        # CLICA VISUALIZAR (na grid de busca)
        clicar_visualizar_primeira_linha(timeout=20)

        # AGORA SIM: “Mostrar Mais → Todos” NA PÁGINA DE DOCUMENTOS
        abrir_dropdown_todos_documentos(timeout=12)

        # EXTRAI DADOS
        info = scrape_prontuario(driver, timeout=15)

        # PREPARA SAÍDA
        nome = info['nome']
        em = info['id']
        cpf = info['cpf']
        total_num = so_digitos(info['total_documentos'])
        docs_raw = info['documentos']

        docs_agregados, mapa_contagens = agrega_docs(docs_raw)
        docs_normalizados_lista = list(mapa_contagens.keys())
        status = status_pesquisa(documento, docs_normalizados_lista)

        # Colunas fixas F..I com “Pendente” se ausentes
        def texto_doc_fixo(rotulo):
            qtd = mapa_contagens.get(rotulo, 0)
            if qtd == 0:
                return "Pendente"
            return f"{rotulo} ({qtd})" if qtd > 1 else rotulo

        col_F = texto_doc_fixo(COL_FIXA_F)
        col_G = texto_doc_fixo(COL_FIXA_G)
        col_H = texto_doc_fixo(COL_FIXA_H)
        col_I = texto_doc_fixo(COL_FIXA_I)

        # Remove dos “extras” os que já entraram nas colunas fixas
        extras = [d for d in docs_agregados
                  if normaliza_titulo(d).split(" (")[0] not in {
                      COL_FIXA_F, COL_FIXA_G, COL_FIXA_H, COL_FIXA_I
                  }]

        garantir_cabecalho_docs(len(extras))
        linha = [nome, em, cpf, total_num, status, col_F, col_G, col_H, col_I] + extras
        ws_out.append(linha)
        wb_out.save(OUT_PATH)

        print("====================================")
        print(f"Nome: {nome}")
        print(f"ID: {em}")
        print(f"CPF: {cpf}")
        print(f"Total de Documentos (n°): {total_num}")
        print(f"Status Doc da Pesquisa: {status}")
        print(f"F: {col_F} | G: {col_G} | H: {col_H} | I: {col_I}")
        if extras:
            print("Extras:", ", ".join(extras))

    except Exception as e:
        print(f"Erro no processamento desse ID {current_id}: {e}")
        try:
            # Linha mínima com pendências nas fixas
            ws_out.append([None, current_id, None, 0, "Não Encontrado",
                           "Pendente", "Pendente", "Pendente", "Pendente"])
            wb_out.save(OUT_PATH)
        except Exception:
            pass
    finally:
        # Volta para a tela de pesquisa
        try:
            voltar_para_pesquisa(timeout=20)
        except Exception:
            try:
                driver.back(); time.sleep(0.8)
                driver.back(); time.sleep(0.8)
                voltar_para_pesquisa(timeout=20)
            except Exception:
                print("Aviso: não consegui voltar à tela de pesquisa automaticamente.")

print(f"✅ Finalizado. Saída em: {OUT_PATH}")
