"""
O principal objetivo desse módulo é facilitar o cadastro de produtos no wordpress, automatizando a etapa de renomeação e
organização das imagens para que então sejam enviadas para o site.
A ideia é que antes que seja realizado o cadastro dos produtos, seja realizado o upload das imagens no wordpress, e com
o link dessas imagens em mãos é possível realizar o cadastro dos produtos já associando as
imagens dos produtos, agilizando todo o processo. É na geração desses links, que estão associados aos nomes das imagens,
que esse modulo se torna útil, manipulando as imagens, renomeando-as, gerando os links e os colocando numa planilha
prontos para serem utilizados.

Algumas configurações, como nomes de colunas, links e arquivos estão localizadas no arquivo constantes.py"""

import os
from pandas import DataFrame
from modulos_externos.salvar_ajustar import (
    gerar_dataframe,
    escolher_arquivo,
    salvar_arquivo_planilha,
    escolher_pasta,
)
from typing import List, Union
from constantes import *


class PlanilhaMapaImagens:
    """Classe que representa a planilha que contém as informações para que sejam realizados os demais processos, apenas
    por meio dessa classe é possível manipular a planilha selecionada.

    Sobre a planilha:
    Deve conter, as colunas NOMES_SUBPASTAS e NOMES_URL (pode ter outras).
    A coluna NOMES_SUBPASTAS deve ter os nomes das subpastas em que estão as
    imagens. Obviamente essas subpastas devem estar dentro de uma pasta que será selecionada pelo usuário.
    A coluna NOMES_URL, deve registrar o nome que será base para que as imagens dentro da respectiva subpasta sejam renomeadas
    , seqguindo um sequêncial de nomebase-01, nomebase-02, nomebase-03...
    """

    def __init__(self, planilha: DataFrame):
        self._planilha = planilha
        self.valida()

    def valida(self):
        ...

    def dados_todas_as_subpastas(self) -> List:
        return self._planilha[NOME_COLUNA_NOMES_SUBPASTAS]

    def dados_todos_os_novos_nomes(self) -> List:
        return self._planilha[NOME_COLUNA_NOVOS_NOMES]

    def dados_contagem_fotos(self) -> List:
        return self._planilha[NOME_COLUNA_CONTAGEM_FOTOS]

    def mapa(self) -> List:
        return self._planilha.to_dict(orient="records")

    def numero_registros(self) -> int:
        return len(self.mapa())

    def adicionar_coluna_contagem_fotos(self, num_imagens: list):
        self._planilha[NOME_COLUNA_CONTAGEM_FOTOS] = num_imagens

    def adicionar_coluna_link_planilha(self, links: list):
        self._planilha[NOME_COLUNA_CONTAGEM_FOTOS] = links

    def salvar(self):
        salvar_arquivo_planilha(self._planilha, NOME_PLANILHA_CONCLUIDA, "xlsx")


class Renomeador:
    """Classe responsável por renomer as imagens e guardar a informação da contagem de fotos, informação que posterior
    mente será utilizada para geração dos links"""

    def __init__(self, dados_nomes_subpastas: List, dados_novos_nomes: List):
        self.dados_nomes_subpastas = dados_nomes_subpastas
        self.dados_novos_nomes = dados_novos_nomes
        self.contagem_fotos = []

    def _guardar_inf_de_contagem(self, contagem):
        self.contagem_fotos.append(contagem)

    def renomear_imagens_em_subpastas(self) -> None:
        """A pasta selecionada não pode ter outras pastas senão aquelas que deseja renomear. Além disso, deve conter
        apenas pastas, e não arquivos."""

        folder_images = escolher_pasta("Selecione o local das imagens originais")
        folder_destino = escolher_pasta("Selecione o local no qual as imagens renomeadas serão salvas")

        input(
            "ATENÇÃO FAÇA UM BACKUP DE SUA PASTA ANTES DE INICIAR O PROCEDIMENTO, AS IMAGENS SERÃO RECORTADAS"
            "DE UMA PASTA PARA OUTRA. PRESS ANYKEY"
        )

        for index, nome_sub in enumerate(self.dados_nomes_subpastas):
            contagem_arquivos = 0
            try:
                for image_name in os.listdir(folder_images + "\\" + nome_sub):
                    contagem_arquivos += 1

                    os.rename(
                        folder_images + "\\" + nome_sub + "\\" + image_name,
                        folder_destino
                        + "\\"
                        + f"{self.dados_novos_nomes[index]}-{contagem_arquivos}"
                        f".{FORMATO_IMAGEM}",
                    )
                    print(f"Arquivo [{image_name}] renomeado!")
            except OSError:
                print(f"Erro com a pasta {nome_sub}")
            self._guardar_inf_de_contagem(contagem_arquivos)


class GeradorLinks:
    """Classe responsável por gerar os links das imagens e por criar uma lista com todos os conjuntos de links
    criados. """

    @staticmethod
    def _gerar_link(url_imagem: str, numero_imagens: int) -> str:
        """Verifica quantas imagens tem e cria o registro como texto para ser adicionado na coluna de imagens"""
        imagens_str = []
        if numero_imagens == 1:
            imagens_str.append(f"{LINK_PROJETO}/{url_imagem}-1.{FORMATO_IMAGEM}")
        else:
            for n in range(1, numero_imagens + 1):
                imagens_str.append(f"{LINK_PROJETO}/{url_imagem}-{n}.{FORMATO_IMAGEM}")
        image_url_separated_by_quotes = "; ".join(imagens_str)
        return image_url_separated_by_quotes

    @staticmethod
    def criar_lista_de_dados_links_imagens(
        todos_os_novos_nomes: List, dados_contagem_fotos: List
    ) -> Union[List, None]:
        """"""
        if len(todos_os_novos_nomes) == len(dados_contagem_fotos):
            dados_coluna = []
            for n in range(0, len(todos_os_novos_nomes)):
                url_imagem = todos_os_novos_nomes[n]
                numero_fotos = dados_contagem_fotos[n]
                dados_coluna.append(GeradorLinks._gerar_link(url_imagem, numero_fotos))
            return dados_coluna
        else:
            print("Os conjuntos de informações não possuem o mesmo tamanho.")
            return


if __name__ == "__main__":
    planilha_mapa = PlanilhaMapaImagens(
        gerar_dataframe(escolher_arquivo("Selecione a planilha mapa"))
    )
    renomeador = Renomeador(
        planilha_mapa.dados_todas_as_subpastas(),
        planilha_mapa.dados_todos_os_novos_nomes(),
    )
    renomeador.renomear_imagens_em_subpastas()
    planilha_mapa.adicionar_coluna_contagem_fotos(renomeador.contagem_fotos)
    dados_links = GeradorLinks.criar_lista_de_dados_links_imagens(
        planilha_mapa.dados_todos_os_novos_nomes(), planilha_mapa.dados_contagem_fotos()
    )
    planilha_mapa.adicionar_coluna_link_planilha(dados_links)
    planilha_mapa.salvar()
