---
title: Solucionar erros de desenvolvimento com suplementos do Office
description: Saiba como solucionar erros de desenvolvimento em suplementos do Office.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: 48216230db4bf90ca53ef10d98786877bd3905c2
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771421"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Solucionar erros de desenvolvimento com suplementos do Office

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Não é possível carregar o suplemento no painel de tarefas ou outros problemas relacionados ao manifesto do suplemento

Confira [Validar o manifesto de suplemento do Office](troubleshoot-manifest.md) e [Depurar seu suplemento com o log do tempo de execução](runtime-logging.md) para depurar problemas de manifesto de suplemento.

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Alterações nos comandos de suplemento, incluindo botões da faixa de opções e itens de menu, não entram em vigor

Se alterações feitas no manifesto, como nomes de arquivo de ícones de botão da faixa de opções ou texto de comandos de suplemento, não parecerem entrar em vigor, experimente limpar o cache do Office no computador. 

#### <a name="for-windows"></a>Para Windows:

Exclua o conteúdo da pasta `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` e exclua o conteúdo da pasta `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` , se ela existir.

#### <a name="for-mac"></a>Para Mac:

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>No iOS:
Chame `window.location.reload(true)` usando o JavaScript no suplemento para forçar um recarregamento. Outra alternativa é reinstalar o Office.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>Alterações em arquivos estáticos, como JavaScript, HTML e CSS, não entram em vigor

O navegador pode estar armazenando esses arquivos em cache. Para evitar isso, desative o cache do lado do cliente ao desenvolver. Os detalhes dependerão do tipo de servidor que você estiver usando. Na maioria dos casos, envolve adicionar determinados cabeçalhos às respostas HTTP. Sugerimos o seguinte conjunto:

- Controle de cache: "privado, sem cache, sem armazenamento"
- Pragma: "sem cache"
- Expira: "-1"

Para um exemplo de como fazer isso em um servidor Node.JS Express, confira [este arquivo app.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js). Para um exemplo em um projeto ASP.NET, confira [este arquivo cshtml](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).

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

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>As alterações feitas nos valores de propriedade não acontecem e não há mensagem de erro

Verifique a documentação de referência da propriedade para ver se ela é somente leitura. Além disso, as [definições do TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) para o Office js especificam quais propriedades de objeto são somente leitura. Se você tentar definir uma propriedade somente leitura, a operação de gravação falhará silenciosamente, sem nenhum erro. O exemplo a seguir tenta erroneamente definir a propriedade somente leitura [Chart.ID](/javascript/api/excel/excel.chart#id). Consulte também [algumas propriedades não podem ser definidas diretamente](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>Obter erro: "este suplemento não está mais disponível"

A seguir estão algumas das causas desse erro. Se você descobrir causas adicionais, diga-nos com a ferramenta de feedback na parte inferior da página.

- Se você estiver usando o Visual Studio, pode haver um problema com o Sideload. Feche todas as instâncias do host do Office e do Visual Studio. Reinicie o Visual Studio e tente pressionar F5 novamente.
- O manifesto do suplemento foi removido de seu local de implantação, como implantação centralizada, um catálogo do SharePoint ou um compartilhamento de rede.
- O valor do elemento [ID](../reference/manifest/id.md) no manifesto foi alterado diretamente na cópia implantada. Se, por qualquer motivo, você quiser alterar essa ID, primeiro remova o suplemento do host do Office e, em seguida, substitua o manifesto original pelo manifesto alterado. Você precisa limpar o cache do Office para remover todos os rastros do original. Consulte a seção [as alterações nos comandos de suplemento, incluindo botões de faixa de opções e itens de menu, não entram em vigor](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) anteriormente neste artigo.
- O manifesto do suplemento tem um `resid` que não está definido em nenhum lugar na seção de [recursos](../reference/manifest/resources.md) do manifesto ou há uma incompatibilidade na ortografia do `resid` local de origem entre onde é usado e onde está definido na `<Resources>` seção.
- Há um `resid` atributo em algum lugar no manifesto com mais de 32 caracteres. Um `resid` atributo, e o `id` atributo do recurso correspondente na `<Resources>` seção, não podem ter mais de 32 caracteres.
- O suplemento tem um comando personalizado do suplemento, mas você está tentando executá-lo em uma plataforma que não dá suporte a eles. Para saber mais, confira [conjuntos de requisitos de comandos de suplemento](../reference/requirement-sets/add-in-commands-requirement-sets.md).

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>O suplemento não funciona na borda, mas funciona em outros navegadores

Consulte [solução de problemas do Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>O suplemento do Excel gera erros, mas não consistentemente

Confira [solucionar problemas de suplementos do Excel](../excel/excel-add-ins-troubleshooting.md) para possíveis causas.

## <a name="see-also"></a>Confira também

- [Depurar suplementos no Office na Web](debug-add-ins-in-office-online.md)
- [Realizar sideload de um suplemento do Office no iPad e no Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Depurar suplementos do Office no iPad e no Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Extensão de depuração de suplementos do Microsoft Office para o Visual Studio Code](debug-with-vs-extension.md)
- [Validar o manifesto de suplemento do Office](troubleshoot-manifest.md)
- [Depurar seu suplemento com o log do tempo de execução](runtime-logging.md)
- [Solucionar erros de usuários com Suplementos do Office](testing-and-troubleshooting.md)
