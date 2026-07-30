# -----------------------------------------------------------------------------
# Encontra nomes que aparecem no ORGANOGRAMA (.pptx) mas NÃO existem na base
# Fênix - o caminho inverso do validacao_organograma_final.py (que checa
# "quem do Fênix está no PPT"). Útil pra achar gente desligada, terceirizada,
# ou nome digitado errado no PPT que não bate com nada no Fênix.
#
# Requisitos: pip install python-pptx pandas openpyxl xlrd
# -----------------------------------------------------------------------------

import pandas as pd
import validacao_organograma_final as base

# === CONFIGURAÇÃO: edite só estes caminhos ==================================
# IMPORTANTE: sempre use r"..." (raw string) em caminhos do Windows.
CAMINHO_PPTX = r"C:\Users\augus\OneDrive\Documentos\Automação\Organograma_Atacado IPC_Jul26.pptx"
CAMINHO_FENIX = r"C:\Users\augus\OneDrive\Documentos\Automação\Base_Fenix.xlsx"
CAMINHO_SAIDA = r"C:\Users\augus\OneDrive\Documentos\Automação\nomes_pptx_sem_fenix.csv"
# =============================================================================


def encontra_no_fenix(nome_norm, fenix_nomes_norm):
    """Mesma lógica de matching do script principal (subsequência ordenada +
    sigla + truncamento), só que na direção contrária: procura esse nome do
    PPT em algum lugar da lista de nomes do Fênix."""
    tk = nome_norm.split()
    if not tk:
        return False
    first = tk[0]
    for fn in fenix_nomes_norm:
        fn_tk = fn.split()
        if fn_tk and fn_tk[0] == first and base.name_tokens_match(tk, fn_tk)[0]:
            return True
    return False


def main():
    fenix = pd.read_excel(CAMINHO_FENIX)
    fenix['nome_norm'] = fenix['Nome'].map(base.norm)
    known_names = pd.concat([fenix['Nome'], fenix['Gestor']]).dropna().unique().tolist()
    fenix_nomes_norm = set(fenix['nome_norm'].dropna())

    pptx = base.extract_names_vinculo(CAMINHO_PPTX, known_names=known_names)
    pptx['nome_norm'] = pptx['nome'].map(base.norm)

    # Junta todas as ocorrências da mesma pessoa (pode aparecer em vários
    # slides) numa linha só, guardando em quais slides ela apareceu.
    agrupado = (
        pptx.groupby('nome_norm')
            .agg(nome=('nome', 'first'),
                 vinculo_pptx=('vinculo_pptx', lambda s: s.dropna().iloc[0] if s.notna().any() else None),
                 slides=('slide', lambda s: ', '.join(str(x) for x in sorted(set(s)))))
            .reset_index()
    )

    resultados = []
    for _, row in agrupado.iterrows():
        encontrado = encontra_no_fenix(row['nome_norm'], fenix_nomes_norm)
        if not encontrado:
            resultados.append(dict(
                nome=row['nome'], vinculo_pptx=row['vinculo_pptx'],
                slides=row['slides'],
            ))

    res = pd.DataFrame(resultados)
    res.to_csv(CAMINHO_SAIDA, index=False, encoding='utf-8-sig')

    print(f"Total de pessoas únicas no PPT: {len(agrupado)}")
    print(f"Não encontradas no Fênix: {len(res)}")
    print()
    if len(res):
        print(res.to_string())
    print()
    print(f"Arquivo salvo em: {CAMINHO_SAIDA}")


if __name__ == "__main__":
    main()
