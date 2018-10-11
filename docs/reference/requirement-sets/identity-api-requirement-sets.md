# <a name="identity-api-requirement-sets"></a>Identificar conjuntos de requisitos de API

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou uma verificação em tempo de execução para determinar se um host do Office oferece suporte às APIs necessárias a um suplemento. Para obter mais informações, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Office são executados em diversas versões do Office. A tabela a seguir lista os conjuntos de requisitos da API de identidade, os aplicativos host do Office que oferecem suporte a esse conjunto de requisitos e os números de verão ou build do aplicativo do Office.

|  Conjunto de requisitos  | Office 2013 para Windows | Office 365 para Windows   |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com e Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | N/D | Versão prévia ***** | Em breve | Versão prévia *****| Versão prévia | Versão prévia| Em breve | Em breve |

> ***** Durante a fase de versão prévia, a API de Identidade tem suporte no Windows 2016 e no Mac apenas para usuários do programa Insiders usando a opção Fast. Para ingressar no programa Insiders, confira [Seja um Insider do Office](https://products.office.com/office-insider?tab=tab-1). Para alternar para o Fast track, confira [Insider Fast](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961).

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- [Números de versão e de build de lançamentos do canal de atualizações para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos comuns da API do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="identityapi-11"></a>IdentityAPI 1.1 

A IdentityAPI 1.1 de Logon Único é a primeira versão da API. Para obter detalhes sobre essa API, confira a seção  [referência da API de SSO](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) em [Habilitar SSO em um suplemento](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos de API e hosts do Office](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
