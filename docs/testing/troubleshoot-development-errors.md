---
title: Solucionar erros de desenvolvimento com Suplementos do Office
description: Saiba como solucionar erros de desenvolvimento em Office de complementos.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2a17a9eafd91cd174209b1974eea61715385c0ad
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990800"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>Solucionar erros de desenvolvimento com Suplementos do Office

Aqui está uma lista de problemas comuns que podem ser encontrados durante o desenvolvimento de um Office Add-in.

> [!TIP]
> Limpar o Office geralmente corrige problemas relacionados ao código desleleo. Isso garante que o manifesto mais recente seja carregado, usando os nomes de arquivo atuais, o texto do menu e outros elementos de comando. Para saber mais, consulte [Clear the Office cache](clear-cache.md).

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>Não é possível carregar o suplemento no painel de tarefas ou outros problemas relacionados ao manifesto do suplemento

Confira [Validar o manifesto de suplemento do Office](troubleshoot-manifest.md) e [Depurar seu suplemento com o log do tempo de execução](runtime-logging.md) para depurar problemas de manifesto de suplemento.

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>Alterações nos comandos de suplemento, incluindo botões da faixa de opções e itens de menu, não entram em vigor

Limpar o cache ajuda a garantir que a versão mais recente do manifesto do seu complemento está sendo usada. Para limpar o cache Office, siga as instruções em [Limpar o Office cache](clear-cache.md). Se você estiver usando Office na Web, limpe o cache do navegador por meio da interface do usuário do navegador.

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>Alterações em arquivos estáticos, como JavaScript, HTML e CSS, não entram em vigor

O navegador pode estar armazenando esses arquivos em cache. Para evitar isso, desative o cache do lado do cliente ao desenvolver. Os detalhes dependerão do tipo de servidor que você estiver usando. Na maioria dos casos, envolve adicionar determinados cabeçalhos às respostas HTTP. Sugerimos o conjunto a seguir.

- Controle de cache: "privado, sem cache, sem armazenamento"
- Pragma: "sem cache"
- Expira: "-1"

Para um exemplo de como fazer isso em um servidor Node.JS Express, confira [este arquivo app.js](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/app.js). Para um exemplo em um projeto ASP.NET, confira [este arquivo cshtml](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).

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

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>Alterações feitas em valores de propriedade não ocorrem e não há mensagem de erro

Verifique a documentação de referência da propriedade para ver se ela é somente leitura. Além disso, as [definições TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) para Office JS especificam quais propriedades de objeto são somente leitura. Se você tentar definir uma propriedade somente leitura, a operação de gravação falhará silenciosamente, sem nenhum erro. O exemplo a seguir tenta definir erroneamente a propriedade somente [leitura](/javascript/api/excel/excel.chart#id)Chart.id . Consulte também [Algumas propriedades não podem ser definidas diretamente](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>Recebendo um erro: "Esse add-in não está mais disponível"

A seguir estão algumas das causas desse erro. Se você descobrir causas adicionais, conte-nos com a ferramenta de comentários na parte inferior da página.

- Se você estiver usando Visual Studio, pode haver um problema com o sideload. Feche todas as instâncias do host Office e Visual Studio. Reinicie Visual Studio e tente pressionar F5 novamente.
- O manifesto do add-in foi removido de seu local de implantação, como Implantação Centralizada, um catálogo SharePoint ou um compartilhamento de rede.
- O valor do elemento [ID](../reference/manifest/id.md) no manifesto foi alterado diretamente na cópia implantada. Se, por qualquer motivo, você quiser alterar essa ID, primeiro remova o complemento do host Office e substitua o manifesto original pelo manifesto alterado. Muitos precisam limpar o cache Office para remover todos os rastreamentos do original. Consulte o [artigo Limpar o Office de cache](clear-cache.md) para obter instruções sobre como limpar o cache do seu sistema operacional.
- O manifesto do add-in tem um que não é definido em qualquer lugar na seção Recursos do manifesto, ou há uma incompatibilidade na ortografia do entre onde ele é usado e onde ele é definido na `resid` [](../reference/manifest/resources.md) `resid` `<Resources>` seção.
- Há um `resid` atributo em algum lugar no manifesto com mais de 32 caracteres. Um `resid` atributo e o atributo do recurso correspondente na seção não podem ter mais de `id` `<Resources>` 32 caracteres.
- O add-in tem um Comando de Complemento personalizado, mas você está tentando executar em uma plataforma que não oferece suporte a eles. Para obter mais informações, consulte [Conjuntos de requisitos de comandos de complemento.](../reference/requirement-sets/add-in-commands-requirement-sets.md)

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>O complemento não funciona no Edge, mas funciona em outros navegadores

Consulte [Solução de problemas Microsoft Edge problemas](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Excel de complemento lança erros, mas não de forma consistente

Consulte [Solução de Excel de soluções para possíveis](../excel/excel-add-ins-troubleshooting.md) causas.

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a>Erros de validação de esquema de manifesto em Visual Studio projetos

Se você estiver usando recursos mais novos que exigem alterações no arquivo de manifesto, poderá obter erros de validação Visual Studio. Por exemplo, ao adicionar o elemento para implementar o tempo de execução `<Runtimes>` javaScript compartilhado, você pode ver o seguinte erro de validação.

**O elemento 'Host' no namespace ' ' tem o elemento filho http://schemas.microsoft.com/office/taskpaneappversionoverrides inválido 'Runtimes' no namespace http://schemas.microsoft.com/office/taskpaneappversionoverrides ' '**

Se isso ocorrer, você poderá atualizar os arquivos XSD que Visual Studio usa para as versões mais recentes. As versões mais recentes do esquema estão [em [MS-OWEMXML]: Apêndice A: Esquema XML completo](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).

### <a name="locate-the-xsd-files"></a>Localizar os arquivos XSD

1. Abra seu projeto em Visual Studio.
1. No **Explorador de Soluções,** abra o arquivo manifest.xml. O manifesto normalmente está no primeiro projeto em sua solução.
1. Escolha **Exibir Janela** de  >  **Propriedades** (F4).
1. Na Janela **Propriedades**, escolha a reellipse (...) para abrir o **editor esquemas XML.** Aqui você pode encontrar o local exato da pasta de todos os arquivos de esquema que seu projeto usa.

### <a name="update-the-xsd-files"></a>Atualizar os arquivos XSD

1. Abra o arquivo XSD que você deseja atualizar em um editor de texto. O nome do esquema do erro de validação será correlacionado ao nome do arquivo XSD. Por exemplo, abra **TaskPaneAppVersionOverridesV1_0.xsd**.
1. Localize o esquema atualizado [em [MS-OWEMXML]: Apêndice A: Esquema XML completo](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8). Por exemplo, TaskPaneAppVersionOverridesV1_0 está no [esquema taskpaneappversionoverrides](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40).
1. Copie o texto para o editor de texto.
1. Salve o arquivo XSD atualizado.
1. Reinicie Visual Studio para buscar as novas alterações de arquivo XSD.

Você pode repetir o processo anterior para quaisquer esquemas adicionais que estão des date.

## <a name="see-also"></a>Confira também

- [Depurar suplementos no Office na Web](debug-add-ins-in-office-online.md)
- [Realizar sideload de um suplemento do Office no iPad e no Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Depurar suplementos do Office em um Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Extensão de Depurador de Suplementos do Microsoft Office para o Visual Studio Code](debug-with-vs-extension.md)
- [Validar o manifesto de suplemento do Office](troubleshoot-manifest.md)
- [Depurar seu suplemento com o log do tempo de execução](runtime-logging.md)
- [Solucionar erros de usuários com Suplementos do Office](testing-and-troubleshooting.md)
- [Microsoft Q&A (office-js-dev)](/answers/topics/office-js-dev.html)
