---
title: Personalizar o suplemento habilitado para SSO do Node.js.
description: Saiba mais sobre como personalizar o suplemento habilitado para SSO que você criou com o gerador Yeoman.
ms.date: 02/20/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: d71206d6b03b8a92e50b316cc75c401866be5334
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608822"
---
# <a name="customize-your-nodejs-sso-enabled-add-in"></a>Personalizar o suplemento habilitado para SSO do Node.js.

> [!IMPORTANT]
> Este artigo cria o suplemento habilitado para SSO que é criado ao concluir o [início rápido de logon único (SSO)](sso-quickstart.md). Conclua o início rápido antes de ler este artigo.

O [início rápido do SSO](sso-quickstart.md) cria um suplemento habilitado para sso que obtém as informações de perfil do usuário conectado e as grava no documento ou na mensagem. Neste artigo, você passará pelo processo de atualização do suplemento que você criou com o gerador Yeoman no início rápido do SSO, para Adicionar nova funcionalidade que exija permissões diferentes.

## <a name="prerequisites"></a>Pré-requisitos

* Um suplemento do Office que você criou seguindo as instruções no [início rápido de SSO](sso-quickstart.md).

* Em pelo menos algumas pastas e arquivos armazenados no OneDrive for Business na assinatura do Office 365.

* [Node.js](https://nodejs.org) (a versão mais recente de [LTS](https://nodejs.org/about/releases)).

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a>Revisar o conteúdo do projeto

Vamos começar com uma revisão rápida do projeto de suplemento [criado anteriormente com o gerador Yeoman](sso-quickstart.md).

> [!NOTE]
> Em lugares onde este artigo faz referência a arquivos de script usando a extensão de arquivo **. js** , considere a extensão de arquivo **. TS** , em vez disso, se o projeto foi criado com TypeScript.

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a>Adicionar nova funcionalidade 

O suplemento que você criou com o início rápido do SSO usa o Microsoft Graph para obter as informações de perfil do usuário conectado e grava essas informações no documento ou na mensagem. Vamos alterar a funcionalidade do suplemento de forma que ele obtenha os nomes dos dez principais arquivos e pastas do OneDrive for Business do usuário conectado e grava essas informações no documento ou na mensagem. Habilitar essa nova funcionalidade requer a atualização das permissões do aplicativo no Azure e a atualização do código no projeto do suplemento.

### <a name="update-app-permissions-in-azure"></a>Atualizar permissões do aplicativo no Azure

Antes que o suplemento possa ler com êxito o conteúdo do OneDrive for Business do usuário, suas informações de registro de aplicativo no Azure devem ser atualizadas com as permissões apropriadas. Complete as etapas a seguir para conceder ao aplicativo a permissão **files. Read. All** e revogar a permissão **User. Read** , que não é mais necessária.

1. Navegue até o [portal do Azure](https://ms.portal.azure.com/#home) e **entre usando suas credenciais de administrador do Office 365**. 

2. Navegue até a página **registros de aplicativos** . 
    > [!TIP]
    > Você pode fazer isso escolhendo o bloco de **registros do aplicativo** na home page do Azure ou usando a caixa de pesquisa na home page para localizar e escolher registros de **aplicativos**.

3. Na página **registros de aplicativos** , escolha o aplicativo que você criou durante o início rápido. 
    > [!TIP]
    > O **nome de exibição** do aplicativo corresponderá ao nome do suplemento que você especificou ao criar o projeto com o gerador Yeoman.

4. Na página Visão geral do aplicativo, escolha **permissões da API** sob o título **gerenciar** no lado esquerdo da página.

5. Na linha **User. Read** da tabela de permissões, escolha as reticências e, em seguida, selecione **revogar consentimento do administrador** no menu exibido.

6. Selecione o botão **Sim, remover** em resposta ao prompt que é exibido.

7. Na linha **User. Read** da tabela Permissions, escolha as reticências e, em seguida, selecione **remover permissão** no menu exibido.

8. Selecione o botão **Sim, remover** em resposta ao prompt que é exibido.

9. Selecione o botão **Adicionar uma permissão** .

10. No painel que é aberto, escolha **Microsoft Graph** e, em seguida, escolha **permissões delegadas**.

11. No painel **solicitar permissões de API** :

    a. Em **arquivos**, selecione **arquivos. Read. All**.

    b. Selecione o botão **adicionar permissões** na parte inferior do painel para salvar essas alterações de permissões.

12. Selecione o botão **conceder consentimento de administrador para [nome do locatário]** .

13. Selecione o botão **Sim** em resposta ao prompt exibido.

### <a name="update-code-in-the-add-in-project"></a>Atualizar código no projeto do suplemento

Para permitir que o suplemento Leia o conteúdo do OneDrive for Business do usuário conectado, você precisará:

- Atualize o código que faz referência à URL, aos parâmetros e ao escopo de acesso necessários do Microsoft Graph.

- Atualize o código que define a interface do usuário do painel de tarefas, para que ele descreva precisamente a nova funcionalidade. 

- Atualize o código que analisa a resposta do Microsoft Graph e o grava no documento ou na mensagem.

As etapas a seguir descrevem essas atualizações.

### <a name="changes-required-for-any-type-of-add-in"></a>Alterações necessárias para qualquer tipo de suplemento

Conclua as seguintes etapas para o seu suplemento, para alterar a URL, os parâmetros e o escopo de acesso do Microsoft Graph e atualizar a interface do usuário do painel de tarefas. Essas etapas são as mesmas, independentemente de qual host do Office seu suplemento está direcionado.

1. Na **./. ENV** arquivo:

    a. Substitua `GRAPH_URL_SEGMENT=/me` pelo seguinte:`GRAPH_URL_SEGMENT=/me/drive/root/children`

    b. Substitua `QUERY_PARAM_SEGMENT=` pelo seguinte:`QUERY_PARAM_SEGMENT=?$select=name&$top=10`

    c. Substitua `SCOPE=User.Read` pelo seguinte:`SCOPE=Files.Read.All`

2. Em **./manifest.xml**, localize a linha `<Scope>User.Read</Scope>` próxima ao final do arquivo e substitua-a pela linha `<Scope>Files.Read.All</Scope>` .

3. Em **./src/Helpers/fallbackauthdialog.js** (ou em **./src/Helpers/fallbackauthdialog.TS** para um projeto TypeScript), localize a cadeia de caracteres `https://graph.microsoft.com/User.Read` e substitua-a pela cadeia de caracteres `https://graph.microsoft.com/Files.Read.All` , tal como `requestObj` é definida da seguinte maneira:

    ```javascript
    var requestObj = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

    ```typescript
    var requestObj: Object = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

4. Em **./src/TaskPane/TaskPane.html**, localize o elemento `<section class="ms-firstrun-instructionstep__header">` e atualize o texto dentro desse elemento para descrever a nova funcionalidade do suplemento.

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

5. Em **./src/TaskPane/TaskPane.html**, localize e substitua as duas ocorrências da cadeia de caracteres `Get My User Profile Information` pela cadeia de caracteres `Read my OneDrive for Business` .

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">Click the <b>Read my OneDrive for Business</b>
            button.</span>
        <div class="clearfix"></div>
    </li>
    ```

    ```html
    <p align="center">
        <button id="getGraphDataButton" class="popupButton ms-Button ms-Button--primary"><span
                class="ms-Button-label">Read my OneDrive for Business</span></button>
    </p>
    ```

6. Em **./src/TaskPane/TaskPane.html**, localize e substitua a cadeia de caracteres `Your user profile information will be displayed in the document.` com a cadeia de caracteres `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.` .

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

7. Atualize o código que analisa a resposta do Microsoft Graph e o grava no documento ou na mensagem, seguindo as orientações na seção que corresponde ao seu tipo de suplemento:

    - [Alterações necessárias para um suplemento do Excel (JavaScript)](#changes-required-for-an-excel-add-in-javascript)
    - [Alterações necessárias para um suplemento do Excel (TypeScript)](#changes-required-for-an-excel-add-in-typescript)
    - [Alterações necessárias para um suplemento do Outlook (JavaScript)](#changes-required-for-an-outlook-add-in-javascript)
    - [Alterações necessárias para um suplemento do Outlook (TypeScript)](#changes-required-for-an-outlook-add-in-typescript)
    - [Alterações necessárias para um suplemento do PowerPoint (JavaScript)](#changes-required-for-a-powerpoint-add-in-javascript)
    - [Alterações necessárias para um suplemento do PowerPoint (TypeScript)](#changes-required-for-a-powerpoint-add-in-typescript)
    - [Alterações necessárias para um suplemento do Word (JavaScript)](#changes-required-for-a-word-add-in-javascript)
    - [Alterações necessárias para um suplemento do Word (TypeScript)](#changes-required-for-a-word-add-in-typescript)

### <a name="changes-required-for-an-excel-add-in-javascript"></a>Alterações necessárias para um suplemento do Excel (JavaScript)

Se o suplemento for um suplemento do Excel que foi criado com JavaScript, faça as seguintes alterações em **./src/Helpers/documentHelper.js**:

1. Encontre a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToExcel(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. Encontre a `filterUserProfileInfo` função e substitua-a pela seguinte função:

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. Encontre a `writeDataToExcel` função e substitua-a pela seguinte função:

    ```javascript
    function writeDataToExcel(result) {
      return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            let innerArray = [];
            innerArray.push(oneDriveInfo[i]);
            data.push(innerArray);
          }
        }

        const rangeAddress = `B5:B${5 + (data.length - 1)}`;
        const range = sheet.getRange(rangeAddress);
        range.values = data;
        range.format.autofitColumns();

        return context.sync();
      });
    }
    ```

4. Exclua a `writeDataToOutlook` função.

5. Exclua a `writeDataToPowerPoint` função.

6. Exclua a `writeDataToWord` função.

Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.

### <a name="changes-required-for-an-excel-add-in-typescript"></a>Alterações necessárias para um suplemento do Excel (TypeScript)

Se o suplemento for um suplemento do Excel que foi criado com TypeScript, abra **./src/TaskPane/TaskPane.TS**, localize a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    }

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        let innerArray = [];
        innerArray.push(itemNames[i]);
        data.push(innerArray);
      }
    }
    
    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
```

Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.

### <a name="changes-required-for-an-outlook-add-in-javascript"></a>Alterações necessárias para um suplemento do Outlook (JavaScript)

Se o suplemento for um suplemento do Outlook que foi criado com JavaScript, faça as seguintes alterações em **./src/Helpers/documentHelper.js**:

1. Encontre a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToOutlook(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to message. " + error.toString()));
        }
      });
    }
    ```

2. Encontre a `filterUserProfileInfo` função e substitua-a pela seguinte função:

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. Encontre a `writeDataToOutlook` função e substitua-a pela seguinte função:

    ```javascript
    function writeDataToOutlook(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
      }

      Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
    }
    ```

4. Exclua a `writeDataToExcel` função.

5. Exclua a `writeDataToPowerPoint` função.

6. Exclua a `writeDataToWord` função.

Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.

### <a name="changes-required-for-an-outlook-add-in-typescript"></a>Alterações necessárias para um suplemento do Outlook (TypeScript)

Se o suplemento for um suplemento do Outlook que foi criado com TypeScript, abra **./src/TaskPane/TaskPane.TS**, localize a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:

```typescript
export function writeDataToOfficeDocument(result: Object): void {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
        if (itemNames[i] !== null) {
        data.push(itemNames[i]);
        }
    }

    let objectNames: string = "";
    for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
    }
    
    Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
}
```

Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.

### <a name="changes-required-for-a-powerpoint-add-in-javascript"></a>Alterações necessárias para um suplemento do PowerPoint (JavaScript)

Se o suplemento for um suplemento do PowerPoint que foi criado com JavaScript, faça as seguintes alterações em **./src/Helpers/documentHelper.js**:

1. Encontre a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToPowerPoint(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. Encontre a `filterUserProfileInfo` função e substitua-a pela seguinte função:

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. Encontre a `writeDataToPowerPoint` função e substitua-a pela seguinte função:

    ```javascript
    function writeDataToPowerPoint(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "\n";
      }

      Office.context.document.setSelectedDataAsync(
        objectNames, 
        function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            throw asyncResult.error.message;
          }
      });
    }
    ```

4. Exclua a `writeDataToExcel` função.

5. Exclua a `writeDataToOutlook` função.

6. Exclua a `writeDataToWord` função.

Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.

### <a name="changes-required-for-a-powerpoint-add-in-typescript"></a>Alterações necessárias para um suplemento do PowerPoint (TypeScript)

Se o suplemento for um suplemento do PowerPoint que foi criado com TypeScript, abra **./src/TaskPane/TaskPane.TS**, localize a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];

  let itemNames: string[] = [];
  let oneDriveItems = result["value"];
  for (let item of oneDriveItems) {
    itemNames.push(item["name"]);
  };

  for (let i = 0; i < itemNames.length; i++) {
    if (itemNames[i] !== null) {
      data.push(itemNames[i]);
    }
  }

  let objectNames: string = "";
  for (let i = 0; i < data.length; i++) {
    objectNames += data[i] + "\n";
  }

  Office.context.document.setSelectedDataAsync(objectNames, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
```

Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.

### <a name="changes-required-for-a-word-add-in-javascript"></a>Alterações necessárias para um suplemento do Word (JavaScript)

Se o suplemento for um suplemento do Word que foi criado com JavaScript, faça as seguintes alterações em **./src/Helpers/documentHelper.js**:

1. Encontre a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToWord(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. Encontre a `filterUserProfileInfo` função e substitua-a pela seguinte função:

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. Encontre a `writeDataToWord` função e substitua-a pela seguinte função:

    ```javascript
    function writeDataToWord(result) {
      return Word.run(function (context) {
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            data.push(oneDriveInfo[i]);
          }
        }

        const documentBody = context.document.body;
        for (let i = 0; i < data.length; i++) {
          if (data[i] !== null) {
            documentBody.insertParagraph(data[i], "End");
          }
        }

        return context.sync();
      });
    }
    ```

4. Exclua a `writeDataToExcel` função.

5. Exclua a `writeDataToOutlook` função.

6. Exclua a `writeDataToPowerPoint` função.

Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.

### <a name="changes-required-for-a-word-add-in-typescript"></a>Alterações necessárias para um suplemento do Word (TypeScript)

Se o suplemento for um suplemento do Word que foi criado com TypeScript, abra **./src/TaskPane/TaskPane.TS**, localize a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Word.run(function(context) {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        data.push(itemNames[i]);
      }
    }

    const documentBody: Word.Body = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}
```

Depois de fazer essas alterações, continue na seção [Experimente](#try-it-out) deste artigo para experimentar o suplemento atualizado em seu site.

## <a name="try-it-out"></a>Experimente

Se o suplemento for um suplemento do Excel, Word ou PowerPoint, conclua as etapas da seção a seguir para experimentá-lo. Se o suplemento for um suplemento do Outlook, conclua as etapas na seção do [Outlook](#outlook) .

### <a name="excel-word-and-powerpoint"></a>Excel, Word e PowerPoint

Execute as etapas a seguir para experimentar um suplemento do Excel, do Word ou do PowerPoint.

1. Na pasta raiz do projeto, execute o seguinte comando para compilar o projeto, inicie o servidor Web local e Sideload seu suplemento no aplicativo cliente do Office selecionado anteriormente.

    > [!NOTE]
    > Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento. Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.

    ```command&nbsp;line
    npm start
    ```

2. No aplicativo cliente do Office que é aberto quando você executa o comando anterior (ou seja, Excel, Word ou PowerPoint), certifique-se de que você está conectado com um usuário que seja membro da mesma organização do Office 365 como a conta de administrador do Office 365 que você usou para se conectar ao Azure durante a [configuração do SSO](sso-quickstart.md#configure-sso) para o aplicativo. Isso estabelecerá as condições apropriadas para que o SSO seja bem-sucedido. 

3. No aplicativo cliente do Office, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento. A imagem a seguir mostra esse botão no Excel.

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-3b.png)

4. Na parte inferior do painel de tarefas, escolha o botão **ler meu onedrive for Business** para iniciar o processo de SSO. 

5. Se uma janela de diálogo for exibida solicitando permissões em nome do suplemento, isso significa que não há suporte ao SSO no seu cenário e, em vez disso, o suplemento voltou para um método alternativo de autenticação do usuário. Isso pode ocorrer quando o administrador do locatário não tiver consentido ao suplemento acesso ao Microsoft Graph, ou quando o usuário não estiver conectado ao Office com uma conta válida da Microsoft ou do Office 365 ("Corporativa ou de Estudante"). Escolha o botão **Aceitar** na janela de diálogo para continuar.

    ![Caixa de diálogo Solicitação de permissões](../images/sso-permissions-request.png)

    > [!NOTE]
    > Após um usuário aceitar a solicitação de permissões, elas não serão solicitadas novamente no futuro.

6. O suplemento lê os dados do OneDrive for Business do usuário conectado e grava os nomes dos 10 arquivos e pastas principais no documento. A imagem a seguir mostra um exemplo de nomes de arquivos e pastas gravados em uma planilha do Excel.

    ![Informações sobre o OneDrive for Business na planilha do Excel](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a>Outlook

Execute as etapas a seguir para experimentar um suplemento do Outlook.

1. Na pasta raiz do projeto, execute o seguinte comando para compilar o projeto e iniciar o servidor Web local.

    > [!NOTE]
    > Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento. Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.

    ```command&nbsp;line
    npm start
    ```

2. Siga as instruções [Realizar sideload dos suplementos do Outlook para teste](/outlook/add-ins/sideload-outlook-add-ins-for-testing)para realizar o sideload do suplemento do Outlook. Certifique-se de que você está conectado ao Outlook com um usuário que é membro da mesma organização do Office 365 que a conta de administrador do Office 365 que você usou para se conectar ao Azure durante a [configuração do SSO](sso-quickstart.md#configure-sso) para o aplicativo. Isso estabelecerá as condições apropriadas para que o SSO seja bem-sucedido. 

3. Escreva uma nova mensagem no Outlook.

4. Na janela redigir mensagem, escolha o botão **Exibir painel de tarefas** na faixa de opções para abrir o painel de tarefas de suplemento.

    ![Botão do suplemento do Outlook](../images/outlook-sso-ribbon-button.png)

5. Na parte inferior do painel de tarefas, escolha o botão **ler meu onedrive for Business** para iniciar o processo de SSO. 

6. Se uma janela de diálogo for exibida solicitando permissões em nome do suplemento, isso significa que não há suporte ao SSO no seu cenário e, em vez disso, o suplemento voltou para um método alternativo de autenticação do usuário. Isso pode ocorrer quando o administrador do locatário não tiver consentido ao suplemento acesso ao Microsoft Graph, ou quando o usuário não estiver conectado ao Office com uma conta válida da Microsoft ou do Office 365 ("Corporativa ou de Estudante"). Escolha o botão **Aceitar** na janela de diálogo para continuar.

    ![Caixa de diálogo Solicitação de permissões](../images/sso-permissions-request.png)

    > [!NOTE]
    > Após um usuário aceitar a solicitação de permissões, elas não serão solicitadas novamente no futuro.

7. O suplemento lê os dados do OneDrive for Business do usuário conectado e grava os nomes dos 10 arquivos e pastas principais no corpo da mensagem de email.

    ![Informações sobre o OneDrive for Business na mensagem do Outlook](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a>Próximas etapas

Parabéns, você personalizou com êxito a funcionalidade do suplemento habilitado para SSO que você criou com o gerador Yeoman no [início rápido de SSO](sso-quickstart.md). Para saber mais sobre as etapas de configuração do SSO que o gerador Yeoman concluiu automaticamente e o código que facilita o processo de SSO, confira o tutorial [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md).

## <a name="see-also"></a>Confira também

- [Habilitar o logon único para Suplementos do Office](../develop/sso-in-office-add-ins.md)
- [Início rápido logon único (SSO).](sso-quickstart.md)
- [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md)
- [Solucionar problemas de mensagens de erro no logon único (SSO)](../develop/troubleshoot-sso-in-office-add-ins.md)
