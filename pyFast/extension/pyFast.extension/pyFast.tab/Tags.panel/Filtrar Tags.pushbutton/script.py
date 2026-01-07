# -*- coding: utf-8 -*-
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from pyrevit import forms
import traceback 

# --- 1. O "OPERÁRIO" (Handler) ---
class FiltroHandler(IExternalEventHandler):
    # Mudança: O Handler agora recebe a 'janela' para poder conversar com ela
    def __init__(self, janela_principal):
        self.janela = janela_principal
        self.texto_busca = ""
        self.regra = ""
        self.cat_escolhida = ""
    
    def Execute(self, uiapp):
        try:
            import Autodesk.Revit.DB as DB
            from System.Collections.Generic import List
            
            # Avisa na janela que começou a trabalhar
            self.janela.txtStatus.Text = "Processando..."
            
            uidoc = uiapp.ActiveUIDocument
            doc = uidoc.Document
            
            # --- COLETA ---
            coletor = DB.FilteredElementCollector(doc, doc.ActiveView.Id)\
                      .OfClass(DB.IndependentTag)\
                      .ToElements()

            elementos_filtrados = []
            busca_num = self.para_numero(self.texto_busca)

            # --- FILTRAGEM ---
            for elemento in coletor:
                # 1. Categoria
                if self.cat_escolhida and self.cat_escolhida != "<Todas as Categorias>":
                    if elemento.Category.Name != self.cat_escolhida:
                        continue 

                # 2. Leitura
                valor_tag = elemento.TagText
                if valor_tag is None: valor_tag = ""
                
                v_txt = valor_tag.lower()
                b_txt = self.texto_busca.lower()
                v_num = self.para_numero(valor_tag)
                
                match = False 

                # 3. Regras
                if self.regra == "Contém":
                    match = b_txt in v_txt
                elif self.regra == "Não Contém":
                    match = b_txt not in v_txt
                elif self.regra == "Igual a":
                    if v_num is not None and busca_num is not None:
                        match = v_num == busca_num
                    else:
                        match = v_txt == b_txt
                elif self.regra == "Inicia com":
                    match = v_txt.startswith(b_txt)
                elif self.regra == "Termina com":
                    match = v_txt.endswith(b_txt)
                elif self.regra == "É Maior que":
                    if v_num is not None and busca_num is not None:
                        match = v_num > busca_num
                    else:
                        match = v_txt > b_txt
                elif self.regra == "É Menor que":
                    if v_num is not None and busca_num is not None:
                        match = v_num < busca_num
                    else:
                        match = v_txt < b_txt
                
                if match:
                    elementos_filtrados.append(elemento.Id)

            # --- RESULTADO FINAL NA TELA ---
            if len(elementos_filtrados) > 0:
                lista_csharp = List[DB.ElementId](elementos_filtrados)
                uidoc.Selection.SetElementIds(lista_csharp)
                uidoc.RefreshActiveView()
                
                # MUDANÇA: Escreve direto na janela XAML em vez de printar
                msg = "Sucesso! {} tags selecionadas.".format(len(elementos_filtrados))
                self.janela.txtStatus.Text = msg
                self.janela.txtStatus.Foreground = self.janela.cor_sucesso # Usa cor verde (definida na janela)
            else:
                uidoc.Selection.SetElementIds(List[DB.ElementId]())
                
                self.janela.txtStatus.Text = "Nenhuma tag encontrada com esse critério."
                self.janela.txtStatus.Foreground = self.janela.cor_erro # Usa cor vermelha

        except Exception as e:
            self.janela.txtStatus.Text = "Erro: " + str(e)
            self.janela.txtStatus.Foreground = self.janela.cor_erro
            traceback.print_exc()

    def GetName(self):
        return "FiltroTagsHandler"

    def para_numero(self, texto):
        try:
            if texto is None: return None
            return float(texto.replace(',', '.'))
        except ValueError:
            return None


# --- 2. A JANELA ---
class JanelaFiltro(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, 'script.xaml')
        
        # Cores para o status (Truque para usar cores do sistema)
        from System.Windows.Media import Brushes
        self.cor_sucesso = Brushes.Green
        self.cor_erro = Brushes.Red
        self.cor_normal = Brushes.Gray

        # Configura o Handler passando "self" (a própria janela)
        self.handler = FiltroHandler(self)
        self.evento_externo = ExternalEvent.Create(self.handler)
        
        try:
            import Autodesk.Revit.DB as DB
            doc = __revit__.ActiveUIDocument.Document
            self.carregar_categorias(doc, DB)
        except:
            pass 

    def carregar_categorias(self, doc, DB):
        tags_vista = DB.FilteredElementCollector(doc, doc.ActiveView.Id)\
                     .OfClass(DB.IndependentTag)\
                     .ToElements()
        cats = set()
        for t in tags_vista:
            if t.Category: cats.add(t.Category.Name)
        
        items = ["<Todas as Categorias>"] + sorted(list(cats))
        self.comboCategoria.ItemsSource = items
        self.comboCategoria.SelectedIndex = 0 

    def btn_clique(self, sender, args):
        # Limpa o status
        self.txtStatus.Text = "Buscando..."
        self.txtStatus.Foreground = self.cor_normal

        # Passa dados
        self.handler.texto_busca = self.inputTexto.Text or ""
        
        if self.comboRegra.SelectedItem:
            self.handler.regra = self.comboRegra.SelectedItem.Content
        else:
            self.handler.regra = "Contém"
            
        self.handler.cat_escolhida = self.comboCategoria.SelectedItem

        # Roda
        self.evento_externo.Raise()

# --- INICIAR ---
JanelaFiltro().Show()