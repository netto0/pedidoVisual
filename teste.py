import re
import unidecode

razoes =["0 | Adriano Braga dos Santos","1 | Antônio Carlos Amaral","2 | Gilvan Delfino de Oliveira",
          "3 | Ilma Delfino de Oliveira",          "4 | João Ferreira dos Santos","5 | José dos Reis Silva",
          "6 | Lindraci Mendes Damascena","7 | Renato Moura Trindade",         "8 | Rodrigues & Pinheiro Ltda",
          "9 | Santos & Souza Refeições Ltda","10| Pérola do Mucuri Sup. Dist. Alim.",
          "11| Terezinha Joaquina Silva & Cia","12| Hélia Lacerda Machado","13| M. dos Santos Pereira & Cia Ltda",
          "14| Maria Cleuza Costa & Cia Ltda","15| Adão Afonso da Silva","16| Prado & Prado Com. de Alim.",
          "17| Débora Gil Alves Silva","18| Eunilto Maia Santos","19| Lidiomar Chaves Resende",
          "20| Lucilene Dias Barbosa",          "21| Reinaldo Ferreira da Silva","22| Roberlan Medeiros",
          "23| Manoelton Santos de Araújo","24| Supermercado Brás Ltda",         "25| Valdívio Leles dos Santos",
          "26| Jaílson de Jesus","27| Jaílton Ferreira dos Santos","28| José Clemente de Jesus",
          "29| Roberto Gomes da Silva","30| Alas da Silva Santos","31| Claudio Souza Cortes",
          "32| Edenaldo Santana Souza",         "33| Ismar Costa Mendes Lima","34| Jorvane Antônio Lima",
          "35| José Afonso Faria","36| Josenélia Farias Lucas",         "37| Maria Emília Silva de Souza",
          "38| Márcio Simão da Silva","39| Mercearia Mineira Ltda","40| Núbia C. B. Leite",
          "41| Oliveira & Leite Ltda","42| Sílvio Cláudio Com. Prod. Alim.","43| Valdomiro Oliveira",
          "44| Adão Brandão dos Santos",          "45| Adilson Ramos Pereira","46| Afrodízio Tenencio de Brito",
          "47| Da Hora e Soares Ltda","48| Fábio de Souza Bom Jardim",          "49| João Souza dos Santos",
          "50| Raquel Pereira Mota","51| Laís Lima Brito","52| Maria da Glória de Jesus Viana","53| Ademir Galdino",
          "54| Davirlei de Jesus Costa","55| Detrez Azevedo Com. Prod. Alim.","56| Edgard Rocha Santos"
          ,"57| Eric Melo de Oliveira","58| Hélcio Ramos Sobral","59| Janilson Soares Araújo",
          "60| José Roberto Dias do Amaral",          "61| Valdirene Bremer Ramalho","62| Milton Alves de Almeida",
          "63| Zilberto Freitas Meireles","64| Fabiano Folgado",          "65| Rosilene L. Espíndola",
          "66| Aline Almeida Lacerda","67| Valdemir Pereira Santos",
]


def to_ascii(ls):
    for i in range(len(ls)):
        ls[i] = unidecode.unidecode(ls[i])


def search(nome, lista):
    to_ascii(lista)
    ref = unidecode.unidecode(nome)
    for l in lista:
        if re.findall(rf'{ref}',l,flags=re.I) != []:
            print(lista.index(l))
        else:
            None


#janela1.Element('-SEARCH-').update(value=f'')
a = 'string'
print(a.find('z'))