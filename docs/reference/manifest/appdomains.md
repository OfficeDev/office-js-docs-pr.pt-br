# <a name="appdomains-element"></a>Elemento AppDomains

Lista qualquer domínio além do domínio especificado no elemento SourceLocation que seu Suplemento do Office utilizará para carregar páginas. Para cada domínio adicional, especifique um elemento AppDomain.

 **Tipo de suplemento:** Conteúdo, Painel de tarefas, E-mail

## <a name="syntax"></a>Sintaxe

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Pode conter

[AppDomain](appdomain.md)

## <a name="remarks"></a>Comentários

Por padrão, o seu suplemento pode carregar qualquer página que esteja no mesmo domínio que o local especificado no elemento **SourceLocation**. Para carregar páginas que não estejam no mesmo domínio que o suplemento, especifique os domínios usando os elementos **AppDomains** e **AppDomain**. Esse elemento não pode estar vazio. 
