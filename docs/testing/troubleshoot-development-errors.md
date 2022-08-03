---
title: Solucionar erros de desenvolvimento com Suplementos do Office
description: Saiba como solucionar problemas de erros de desenvolvimento em Suplementos do Office.
ms.date: 07/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: 18236787ad6ffa9139eb95299723c8935d584668
ms.sourcegitcommit: 143ab022c9ff6ba65bf20b34b5b3a5836d36744c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/03/2022
ms.locfileid: "67177662"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Solucionar erros de desenvolvimento com Suplementos do Office

Aqui está uma lista de problemas comuns que você pode encontrar ao desenvolver um Suplemento do Office.

> [!TIP]
> Limpar o cache do Office geralmente corrige problemas relacionados ao código obsoleto. Isso garante que o manifesto mais recente seja carregado, usando os nomes de arquivo atuais, o texto do menu e outros elementos de comando. Para saber mais, confira [Limpar o cache do Office](clear-cache.md).

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Não é possível carregar o suplemento no painel de tarefas ou outros problemas relacionados ao manifesto do suplemento

Confira [Validar o manifesto de suplemento do Office](troubleshoot-manifest.md) e [Depurar seu suplemento com o log do tempo de execução](runtime-logging.md) para depurar problemas de manifesto de suplemento.

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Alterações nos comandos de suplemento, incluindo botões da faixa de opções e itens de menu, não entram em vigor

Limpar o cache ajuda a garantir que a versão mais recente do manifesto do suplemento esteja sendo usada. Para limpar o cache do Office, siga as instruções em [Limpar o cache do Office](clear-cache.md). Se você estiver usando Office na Web, limpe o cache do navegador por meio da interface do usuário do navegador.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>Alterações em arquivos estáticos, como JavaScript, HTML e CSS, não entram em vigor

O navegador pode estar armazenando esses arquivos em cache. Para evitar isso, desative o cache do lado do cliente ao desenvolver. Os detalhes dependerão do tipo de servidor que você estiver usando. Na maioria dos casos, envolve adicionar determinados cabeçalhos às respostas HTTP. Sugerimos o conjunto a seguir.

- Controle de cache: "privado, sem cache, sem armazenamento"
- Pragma: "sem cache"
- Expira: "-1"

Para um exemplo de como fazer isso em um servidor Node.JS Express, confira [este arquivo app.js](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/app.js). Para um exemplo em um projeto ASP.NET, confira [este arquivo cshtml](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).

Se o seu suplemento estiver hospedado no Servidor de Informações da Internet (IIS), você também poderá adicionar o seguinte ao web.config.

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

Se essas etapas não parecerem funcionar a princípio, talvez seja necessário limpar o cache do navegador. Faça isso através da interface do usuário do navegador. Às vezes, o cache do Microsoft Edge não é limpo com êxito quando você tenta limpá-lo na interface do usuário do Edge. Se isso acontecer, execute o seguinte comando em um prompt de comando do Windows.

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>As alterações feitas nos valores de propriedade não acontecem e não há nenhuma mensagem de erro

Verifique a documentação de referência da propriedade para ver se ela é somente leitura. Além disso, as [definições de TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) para Office JS especificam quais propriedades de objeto são somente leitura. Se você tentar definir uma propriedade somente leitura, a operação de gravação falhará silenciosamente, sem nenhum erro gerado. O exemplo a seguir tenta definir erroneamente a propriedade somente [leitura](/javascript/api/excel/excel.chart#excel-excel-chart-id-member) Chart.id. Consulte também [Algumas propriedades não podem ser definidas diretamente](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>Recebendo o erro: "Este suplemento não está mais disponível"

A seguir estão algumas das causas desse erro. Se você descobrir causas adicionais, informe-nos com a ferramenta de comentários na parte inferior da página.

- Se você estiver usando o Visual Studio, poderá haver um problema com o sideload. Feche todas as instâncias do host do Office e do Visual Studio. Reinicie o Visual Studio e tente pressionar F5 novamente.
- O manifesto do suplemento foi removido de seu local de implantação, como Implantação Centralizada, um catálogo do SharePoint ou um compartilhamento de rede.
- O valor do elemento [de ID](/javascript/api/manifest/id) no manifesto foi alterado diretamente na cópia implantada. Se, por algum motivo, você quiser alterar essa ID, primeiro remova o suplemento do host do Office e substitua o manifesto original pelo manifesto alterado. Muitos precisam limpar o cache do Office para remover todos os rastreamentos do original. Consulte o [artigo Limpar o cache do Office](clear-cache.md) para obter instruções sobre como limpar o cache para seu sistema operacional.
- O manifesto do suplemento tem um que não está definido em qualquer lugar na seção Recursos [](/javascript/api/manifest/resources) do manifesto ou há uma incompatibilidade na ortografia entre o local em que ele é usado e onde ele é definido na seção.`resid` `resid` **\<Resources\>**
- Há um atributo `resid` em algum lugar no manifesto com mais de 32 caracteres. Um `resid` atributo e o atributo `id` do recurso correspondente na **\<Resources\>** seção não podem ter mais de 32 caracteres.
- O suplemento tem um comando de suplemento personalizado, mas você está tentando executar em uma plataforma que não dá suporte a eles. Para obter mais informações, consulte [conjuntos de requisitos de comandos de suplemento](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets).

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>O suplemento não funciona no Edge, mas funciona em outros navegadores

Consulte [Solução de problemas do Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshoot-microsoft-edge-issues).

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>O suplemento do Excel gera erros, mas não consistentemente

Consulte [Solucionar problemas de suplementos do Excel para possíveis](../excel/excel-add-ins-troubleshooting.md) causas.

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a>Erros de validação de esquema de manifesto em projetos do Visual Studio

Se você estiver usando recursos mais recentes que exigem alterações no arquivo de manifesto, poderá receber erros de validação no Visual Studio. Por exemplo, ao adicionar o elemento **\<Runtimes\>** para implementar o runtime de JavaScript compartilhado, você poderá ver o seguinte erro de validação.

**O elemento 'Host' no namespace 'http://schemas.microsoft.com/office/taskpaneappversionoverrides' tem o elemento filho inválido 'Runtimes' no namespace 'http://schemas.microsoft.com/office/taskpaneappversionoverrides'**

Se isso ocorrer, você poderá atualizar os arquivos XSD que o Visual Studio usa para as versões mais recentes. As versões mais recentes do esquema estão [no [MS-OWEMXML]: Apêndice A: Esquema XML completo](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).

### <a name="locate-the-xsd-files"></a>Localizar os arquivos XSD

1. Abra seu projeto no Visual Studio.
1. No **Gerenciador de Soluções**, abra o manifest.xml arquivo. O manifesto normalmente está no primeiro projeto em sua solução.
1. Escolha **a janela Exibir** > **Propriedades** (F4).
1. Na Janela **Propriedades**, escolha as reticências (...) para abrir o editor de **Esquemas XML** . Aqui você pode encontrar o local exato da pasta de todos os arquivos de esquema que seu projeto usa.

### <a name="update-the-xsd-files"></a>Atualizar os arquivos XSD

1. Abra o arquivo XSD que você deseja atualizar em um editor de texto. O nome do esquema do erro de validação será correlacionado ao nome do arquivo XSD. Por exemplo, abra **TaskPaneAppVersionOverridesV1_0.xsd**.
1. Localize o esquema atualizado [em [MS-OWEMXML]: Apêndice A: Esquema XML completo](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). Por exemplo, TaskPaneAppVersionOverridesV1_0 está no [esquema taskpaneappversionoverrides](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40).
1. Copie o texto para o editor de texto.
1. Salve o arquivo XSD atualizado.
1. Reinicie o Visual Studio para selecionar as novas alterações de arquivo XSD.

Você pode repetir o processo anterior para quaisquer esquemas adicionais que estejam desatualizados.

## <a name="when-working-offline-no-office-apis-work"></a>Ao trabalhar offline, nenhuma APIs do Office funciona

Quando você estiver carregando a Biblioteca JavaScript do Office de uma cópia local em vez da CDN, as APIs poderão parar de funcionar se a biblioteca não estiver atualizada. Se você estiver ausente de um projeto há algum tempo, reinstale a biblioteca para obter a versão mais recente. O processo varia de acordo com o IDE. Escolha uma das opções a seguir com base em seu ambiente.

- **Visual Studio**: consulte [Atualizar para a biblioteca mais recente da API JavaScript do Office](../develop/update-your-javascript-api-for-office-and-manifest-schema-version.md). 
- **Qualquer outro IDE**: consulte os pacotes npm [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) [e @types/office-js](https://www.npmjs.com/package/@types/office-js).

## <a name="see-also"></a>Confira também

- [Depurar suplementos no Office na Web](debug-add-ins-in-office-online.md)
- [Realizar sideload de um Suplemento do Office no Mac](sideload-an-office-add-in-on-mac.md)  
- [Realizar sideload de um Suplemento do Office no iPad](sideload-an-office-add-in-on-ipad.md)  
- [Depurar Suplementos do Office em um Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code](debug-with-vs-extension.md)
- [Validar o manifesto de suplemento do Office](troubleshoot-manifest.md)
- [Depurar seu suplemento com o log do tempo de execução](runtime-logging.md)
- [Solucionar erros de usuários com Suplementos do Office](testing-and-troubleshooting.md)
- [Microsoft Q&A (office-js-dev)](/answers/topics/office-js-dev.html)
