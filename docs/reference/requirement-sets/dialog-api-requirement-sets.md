# <a name="dialog-api-requirement-sets"></a>Conjuntos de requisitos da API de caixa de diálogo

Os conjuntos de requisitos são grupos nomeados de membros da API. Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou uma verificação em tempo de execução para determinar se um host do Office oferece suporte às APIs necessárias a um suplemento. Para obter mais informações, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Office são executados em várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API de caixa de diálogo, os aplicativos host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou de build dos aplicativos do Office.

|  Conjunto de requisitos  | Office 2013 para Windows | Office 2016 para Windows (Instalações MSI)   | Office 365 para Windows (Instalações C2R)   |  Office para iPad  |  Office 365 para Mac  | Office Online  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Build 15.0.4855.1000 ou posterior | Build 16.0.4390.1000 ou posterior | Versão 1602 (Build 6741.0000) ou posterior | 1.22 ou posterior | 15.20 ou posterior| Janeiro de 2017 | Versão 1608 (Build 7601.6800) ou posterior|

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- [Números de versão e de build de lançamentos do canal de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comuns da API do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11"></a>API de caixa de diálogo 1.1 

A API de caixa de diálogo 1.1 é a primeira versão da API. Para obter detalhes sobre a API, consulte o tópico de referência [API de diálogo ](/javascript/api/office/office.ui) .

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos de API e hosts do Office](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
