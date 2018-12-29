---
title: Instale a última versão do Office
description: Informações sobre como desativar essa opção para obter as versões mais recentes do Office.
ms.date: 12/04/2017
ms.openlocfilehash: 96a9a44687ece98c08651e4a4eb2ab29dd3ee6ad
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457583"
---
# <a name="install-the-latest-version-of-office"></a>Instale a última versão do Office

Novos recursos de desenvolvedor, inclusive os que ainda estão na visualização, são fornecidos primeiro aos assinantes que aceitam obter as últimas compilações do Office. 

## <a name="opt-in-to-getting-the-latest-builds"></a>Aceitar para receber as versões mais recentes

Aceitar para receber as versões mais recentes do Office: 

- Se você é assinante do Office 365 Home, Personal ou University, confira [Ser um Office Insider](https://products.office.com/office-insider).
- Se você for um cliente corporativo do Office 365, confira [Instalar a versão de Primeiro Lançamento do Office 365 para clientes corporativos](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).
- Se você estiver executando o Office em um Mac:
    - Inicie um programa do Office para Mac.
    - Selecione **Verificar Atualizações** no menu Ajuda.
    - Na caixa Microsoft AutoUpdate, marque a caixa para participar do programa Office Insider. 

## <a name="get-the-latest-build"></a>Obtenha a versão mais recente:

Para receber as versões mais recentes do Office: 

1. Baixar a [Ferramenta de Implantação do Office](https://www.microsoft.com/download/details.aspx?id=49117). 
2. Execute a ferramenta. Isso extrai estes dois arquivos: Setup.exe e configuration.xml.
3. Substitua o arquivo configuration.xml pelo [Arquivo de Configuração do Primeiro Lançamento](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. Execute o seguinte comando como administrador: `setup.exe /configure configuration.xml` 

    > [!NOTE]
    > O comando pode demorar muito para ser executado sem indicar o progresso.

Quando o processo de instalação for concluído, você terá os últimos aplicativos do Office instalados. Para verificar se você tem a última compilação, vá até **arquivo** > **conta** em qualquer aplicativo do Office. Em Atualizações do Office, você verá o rótulo (Office Insiders) acima do número de versão.

![Uma captura de tela que mostra informações do produto com o rótulo Office Insiders](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Builds mínimos do Office para conjuntos de requisitos de API JavaScript para Office

Para saber mais sobre os builds mínimos de produtos para cada plataforma dos conjuntos de requisitos de API, confira o seguinte:

- [Conjuntos de requisitos da API JavaScript do Word](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets)
- [Conjuntos de requisitos da API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets)
- [Conjuntos de requisitos da API JavaScript do OneNote](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets)
- [Conjuntos de requisitos da API de Caixa de Diálogo](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)
- [Conjuntos de requisitos da API comum do Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
