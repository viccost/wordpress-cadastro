import os
from pandas import DataFrame
from modulos_externos.salvar_ajustar import gerar_dataframe, escolher_arquivo, salvar_arquivo_planilha, escolher_pasta
from typing import List

nome_coluna_num_fotos = "Numero Fotos"
nome_coluna_nome = "NOME"
nome_coluna_url = "URL"
nome_coluna_fotos = "Fotos"


def renomear_varios_arquivos_subpastas_e_alterar_nomes() -> DataFrame:
    """Utilizado para renomear arquivos conforme nome de uma planilha. A planilha deve conter, na coluna NOME,
    o nome das pastas que estão contidas dentro da folder_images selecionadas. Na coluna URL, deve registrar o nome
    para qual as imagens serão renomeadas. O formato das imagens pode ser informado através da variável image_format, e
    não deve estar preenchido no nome da planilha. É importante ter um backup de suas imagens antes de iniciar o proces.
    Uma nova coluna será gerada na planilha destino, destrinchando as URLS das imagens.

    1º Seleção Pasta de Images
    2º Seleção Arquivo XLSX.
    3º Pasta Destino para salvar
    5º Pasta Destino para salvar planilha
    """
    # exemplo com subpastas, não pode haver outras pastas senão aquelas que deseja renomear. Além disso, deve conter
    # apenas pastas, e não arquivos.

    folder_images = escolher_pasta()
    planilha = gerar_dataframe(escolher_arquivo())
    folder_destino = escolher_pasta()
    num_imagens = []
    image_format = 'jpg'
    pastas_listadas_na_planilha = planilha[nome_coluna_nome]
    nomes_url_imagens = planilha[nome_coluna_url]

    input("ATENÇÃO FAÇA UM BACKUP DE SUA PASTA ANTES DE INICIAR O PROCEDIMENTO, AS IMAGENS SERÃO RECORTADAS"
          "DE UMA PASTA PARA OUTRA. PRESS ANYKEY")


    for index, folder_name in enumerate(pastas_listadas_na_planilha):  # identificando todas as pastas no PATH
        file_number = 1
        for image_name in os.listdir(folder_images + '\\' + folder_name):
            os.rename(folder_images + '\\' + folder_name + '\\' + image_name,
                      folder_destino + '\\' + folder_name + '\\' + f'{nomes_url_imagens[index]}-{file_number}'
                                                                   f'.{image_format}')

            print(f"Arquivo [{image_name}] renomeado!")
            file_number += 1
        num_imagens.append(file_number)
    planilha[nome_coluna_num_fotos] = num_imagens
    return planilha


def criar_lista_de_dados_coluna_imagens(planilha: DataFrame) -> List:
    """"""
    dados_coluna = []

    def criar_dados(url_imagem: str, num_imagens: int) -> str:
        """Verifica quantas imagens tem e cria o registro como texto para ser adicionado na coluna de imagens"""
        imagens_str = []
        num_imagens = int(num_imagens)
        if num_imagens == 1:
            imagens_str.append(
                f"https://projetos.oxecommerce.com.br/ikastore/wp-content/uploads/2022/08/{url_imagem}-1.jpg")
        else:
            for n in range(1, num_imagens+1):
                imagens_str.append(f"https://projetos.oxecommerce.com.br/ikastore/wp-content/uploads/2022/08/{url_imagem}-{n}.jpg")
        image_url_separated_by_quotes = '; '.join(imagens_str)
        return image_url_separated_by_quotes

    for register, row in planilha.iterrows():
        url_imagem = row[nome_coluna_url]
        numero_fotos = row[nome_coluna_num_fotos]
        dados_coluna.append(criar_dados(url_imagem, numero_fotos))
    return dados_coluna


if __name__ == "__main__":
    mapa_imagens = renomear_varios_arquivos_subpastas_e_alterar_nomes()
    # mapa_imagens = gerar_dataframe(escolher_arquivo())
    coluna_imagens = criar_lista_de_dados_coluna_imagens(mapa_imagens)
    mapa_imagens['Imagens'] = coluna_imagens
    folder_destino = escolher_pasta()
    salvar_arquivo_planilha(mapa_imagens, "Imagens Mapeadas", "xlsx", folder_destino)
