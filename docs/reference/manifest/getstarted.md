# <a name="getstarted-element"></a>Elemento GetStarted

Fornece informações usadas pelo texto explicativo que aparece quando o suplemento está instalado em hosts do Word, do Excel, do PowerPoint e do OneNote. O elemento **GetStarted** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).

## <a name="child-elements"></a>Elementos filho

| Elemento                       | Obrigatório | Descrição                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | Sim      | Define onde um suplemento expõe a funcionalidade.     |
| [Descrição](#description)   | Sim      | Uma URL para um arquivo que contém funções JavaScript.|
| [LearnMoreUrl](#learnmoreurl) | Não       | Uma URL para uma página que explica o suplemento em detalhes.   |

### <a name="title"></a>Title 

Obrigatório. O título usado para o início do texto explicativo. O atributo **resid** faz referência a uma identificação válida no elemento **ShortStrings** na seção [Resources](resources.md).

### <a name="description"></a>Descrição

Obrigatório. A descrição / o conteúdo do corpo para o texto explicativo. O atributo **resid** faz referência a uma identificação válida no elemento **LongStrings** na seção [Resources](resources.md).

### <a name="learnmoreurl"></a>LearnMoreUrl

Obrigatório. A URL para uma página onde o usuário pode saber mais sobre o suplemento. O atributo **resid** faz referência a uma identificação válida no elemento **Urls** na seção [Resources](resources.md).

> [!NOTE]
> **LearnMoreUrl** atualmente não é renderizado em clientes do Word, Excel ou PowerPoint. Recomendamos que você adicione essa URL para todos os clientes para que a URL seja processada quando ficar disponível. 

## <a name="see-also"></a>Confira também

Os exemplos de código a seguir utilizam o elemento **GetStarted** :

* [Suplemento web do Excel para manipular formatação de tabelas e gráficos](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [Suplemento do Word JavaScript SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Inserir gráficos do Excel usando o Microsoft Graph em um suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
