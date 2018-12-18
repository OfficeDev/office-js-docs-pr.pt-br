# <a name="appdomain-element"></a>Elemento AppDomain

Especifica um domínio adicional que será usado para carregar páginas na janela do suplemento.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> O valor do elemento **AppDomain** deve incluir o protocolo (ex., `<AppDomain>https://myappdomain<AppDomain>`).

## <a name="contained-in"></a>Contido em

[AppDomains](appdomains.md)

## <a name="remarks"></a>Comentários

Os elementos **AppDomain** deve ser usado para especificar os domínios adicionais diferentes daqueles especificados no elemento [SourceLocation](sourcelocation.md). Confira mais informações em [Manifesto XML de Suplementos do Office](/office/dev/add-ins/develop/add-in-manifests).
