---
title: Atualizar para a biblioteca Office api JavaScript mais recente e o esquema de manifesto de manifesto do complemento versão 1.1
description: Atualize seus arquivos de JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto de suplemento usados no seu projeto de Suplemento do Office para a versão 1.1.
ms.date: 01/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5466b010cb0364d78819942f0a1dcc941e1c1269
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63742923"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a>Atualizar para a biblioteca Office api JavaScript mais recente e o esquema de manifesto de manifesto do complemento versão 1.1

Este artigo descreve como atualizar os arquivos do JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação do manifesto do suplemento no projeto do suplemento do Office para a versão 1.1.

> [!NOTE]
> Os projetos criados Visual Studio 2019 já usarão a versão 1.1. No entanto, há atualizações secundárias ocasionais para a versão 1.1 que você pode aplicar ao usar as técnicas neste artigo.

## <a name="use-the-most-up-to-date-project-files"></a>Usar os arquivos de projeto mais atualizados

Se você usar o Visual Studio para desenvolver seu add-in, para usar os membros mais novos da API JavaScript do Office e os recursos [v1.1](../develop/add-in-manifests.md) do manifesto do complemento (que é validado em relação ao offappmanifest-1.1.xsd), você precisará baixar o Visual Studio 2019. Para baixar Visual Studio 2019, consulte a [página Visual Studio IDE.](https://visualstudio.microsoft.com/vs/) Durante a instalação, você precisará selecionar a carga de trabalho de desenvolvimento do Office/SharePoint.

Se você usar um editor de texto ou um IDE diferente do Visual Studio para desenvolver seu complemento, será necessário atualizar as referências à rede de distribuição de conteúdo (CDN) para Office.js e a versão do esquema referenciada no manifesto do seu complemento.

Para executar um complemento desenvolvido usando novos e atualizados recursos de manifesto de API e de complemento do Office.js, seus clientes devem estar executando o Office 2013 SP1 ou produtos locais de versão posterior e, quando aplicável, o SharePoint Server 2013 SP1 e produtos de servidor relacionados, o Exchange Server 2013 Service Pack 1 (SP1) ou os produtos online equivalentes hospedados: Microsoft 365, SharePoint Online e Exchange Online.

Para baixar os produtos do Office, SharePoint e Exchange SP1, consulte o seguinte:

- [Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft Office 2013 e produtos da área de trabalho relacionados](https://support.microsoft.com/kb/2850036)

- [Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft SharePoint Server 2013 e produtos do servidor relacionados](https://support.microsoft.com/kb/2850035)

- [Descrição do Exchange Server 2013 Service Pack 1](https://support.microsoft.com/kb/2926248)

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Atualização de um projeto de suplemento do Office criado com o Visual Studio

Para projetos criados antes da versão v1.1 da API JavaScript do Office e do esquema de manifesto do complemento, você pode atualizar os arquivos de um projeto usando **o NuGet Gerenciador de Pacotes** e atualizar as páginas HTML do seu complemento para fazer referência a eles.

Observe que o processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a>Atualize os Office de biblioteca da API JavaScript em seu projeto para a versão mais recente

As etapas a seguir atualizarão seus arquivos Office.js biblioteca para a versão mais recente. As etapas usam Visual Studio 2019, mas são semelhantes para versões anteriores Visual Studio.

1. No Visual Studio 2019, abra ou crie um novo projeto de Office de **complemento**.
2. Escolha **Ferramentas** >  **NuGet Gerenciador de Pacotes** >  **Manage Nuget Packages for Solution**.
3. Escolha a guia **Atualizações**.
4. Selecione Microsoft.Office.js. Verifique se a origem do pacote **é nuget.org**.
5. No painel esquerdo, escolha **Instalar e** concluir o processo de atualização do pacote.

Você precisará realizar algumas etapas adicionais para concluir a atualização. Na marca  de cabeça das páginas HTML do seu complemento, comente ou exclua quaisquer referências de script office.js existentes e consulte a biblioteca de API JavaScript do Office atualizada da seguinte forma:

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE]
   > O `/1/` em `office.js` na URL de CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Atualizar o arquivo de manifesto no projeto para usar a versão 1.1 do esquema

No arquivo de manifesto do suplemento, atualize o atributo **xmlns** do elemento **OfficeApp** alterando o valor de versão para `1.1` (mantendo inalterados os atributos diferentes de **xmlns**).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> Depois de atualizar a versão do esquema de manifesto do complemento para 1.1, você precisará remover os elementos **Recursos** e Funcionalidades e substituí-los por elementos [Hosts](../reference/manifest/hosts.md) e [Host](../reference/manifest/host.md) ou os elementos [Requisitos e Requisitos](specify-office-hosts-and-api-requirements.md).

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>Atualização de um projeto de suplemento do Office criado com um editor de texto ou outro IDE

Para projetos criados antes da versão v1.1 da API JavaScript do Office e do esquema de manifesto do complemento, você precisa atualizar as páginas HTML do seu complemento para fazer referência CDN da biblioteca v1.1 e atualizar o arquivo de manifesto do seu complemento para usar o esquema v1.1.

O processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.

Você não precisa de cópias locais dos arquivos da API JavaScript do Office (arquivos Office.js e .js específicos do aplicativo) para desenvolver um Add-in doOffice (fazendo referência ao CDN para Office.js baixa os arquivos necessários em tempo de execução), mas se você quiser uma cópia local dos arquivos de biblioteca, poderá usar o [Utilitário NuGet Command-Line](https://docs.nuget.org/consume/installing-nuget) e o comando para `Install-Package Microsoft.Office.js` baixá-los.

> [!NOTE]
> Para obter uma cópia da XSD (Definição de esquema XML) para o manifesto do suplemento v1.1, confira a listagem em [Referência de esquema para manifestos de Suplementos do Office (v1.1)](../develop/add-in-manifests.md).

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a>Atualize os Office de biblioteca da API JavaScript em seu projeto para usar a versão mais recente

1. Abra as páginas HTML do suplemento no editor de texto ou IDE.

2. Na marca  de cabeça das páginas HTML do seu complemento, comente ou exclua quaisquer referências de script office.js existentes e consulte a biblioteca de API JavaScript do Office atualizada da seguinte forma:

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > O `/1/` na frente de `office.js` na URL de CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Atualizar o arquivo de manifesto no projeto para usar a versão 1.1 do esquema

No arquivo de manifesto do suplemento, atualize o atributo **xmlns** do elemento **OfficeApp** alterando o valor de versão para `1.1` (mantendo inalterados os atributos diferentes de **xmlns**).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> Depois de atualizar a versão do esquema de manifesto do complemento para 1.1, você precisará remover os elementos **Recursos** e Funcionalidades e substituí-los por elementos [Hosts](../reference/manifest/hosts.md) e [Host](../reference/manifest/host.md) ou os elementos [Requisitos e Requisitos](specify-office-hosts-and-api-requirements.md).

## <a name="see-also"></a>Confira também

- [Especificar Office aplicativos e requisitos de API](specify-office-hosts-and-api-requirements.md) ]
- [Entendendo a API de JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript para Office](../reference/javascript-api-for-office.md)
- [Referência de esquema para manifestos de suplementos do Office (versão 1.1)](../develop/add-in-manifests.md)
