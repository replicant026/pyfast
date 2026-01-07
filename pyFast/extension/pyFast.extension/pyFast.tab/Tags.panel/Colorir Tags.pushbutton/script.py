# -*- coding: utf-8 -*-
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from pyrevit import forms
import traceback

# --- 1. O OPERÁRIO (Handler) ---
class ColorirHandler(IExternalEventHandler):
    def __init__(self, janela_principal):
        self.janela = janela_principal
        self.nome_parametro = "ID do comando"
        self.modo = "APLICAR" 
    
    def Execute(self, uiapp):
        try:
            # Importações Locais
            import Autodesk.Revit.DB as DB
            
            self.janela.txtStatus.Text = "Processando..."
            
            uidoc = uiapp.ActiveUIDocument
            doc = uidoc.Document
            view_ativa = doc.ActiveView

            # --- COLETAR TAGS ---
            tags_vista = DB.FilteredElementCollector(doc, view_ativa.Id)\
                         .OfClass(DB.IndependentTag)\
                         .ToElements()

            mapa_cores = {}
            count_modificados = 0
            count_sem_param = 0
            
            t = DB.Transaction(doc, "Colorir Tags: " + self.modo)
            t.Start()

            # --- MODO LIMPEZA ---
            if self.modo == "LIMPAR":
                config_vazia = DB.OverrideGraphicSettings()
                for tag in tags_vista:
                    view_ativa.SetElementOverrides(tag.Id, config_vazia)
                
                t.Commit()
                self.janela.txtStatus.Text = "Pronto! Todas as cores foram removidas."
                self.janela.txtStatus.Foreground = self.janela.cor_sucesso
                return

            # --- MODO COLORIR ---
            
            # 1. Busca Hachura Sólida
            padrao_solido = DB.FillPatternElement.GetFillPatternElementByName(doc, DB.FillPatternTarget.Drafting, "<Solid fill>")
            if not padrao_solido:
                padrao_solido = DB.FillPatternElement.GetFillPatternElementByName(doc, DB.FillPatternTarget.Drafting, "<Preenchimento sólido>")
            
            if not padrao_solido:
                filt_pat = DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement)
                for fp in filt_pat:
                    if fp.GetFillPattern().IsSolid:
                        padrao_solido = fp
                        break

            if not padrao_solido:
                forms.alert("Erro: Não foi possível encontrar o padrão de hachura Sólida.")
                t.RollBack()
                return

            # 2. Iteração
            for tag in tags_vista:
                raw_ids = tag.GetTaggedLocalElementIds()
                elementos_tagged = list(raw_ids) 
                
                if not elementos_tagged:
                    continue 

                id_elemento = elementos_tagged[0]
                elemento = doc.GetElement(id_elemento)

                if not elemento: continue

                # 3. Ler Parâmetro
                param = elemento.LookupParameter(self.nome_parametro)
                
                if not param: param = elemento.LookupParameter("Mark")
                if not param: param = elemento.LookupParameter("Marca")
                if not param: param = elemento.LookupParameter("Comments")
                if not param: param = elemento.LookupParameter("Comentários")

                valor_chave = ""
                if param and param.HasValue:
                    valor_chave = param.AsString()
                    if not valor_chave and param.StorageType == DB.StorageType.Double:
                        valor_chave = str(round(param.AsDouble(), 2))
                    elif not valor_chave and param.StorageType == DB.StorageType.Integer:
                        valor_chave = str(param.AsInteger())
                
                if not valor_chave:
                    count_sem_param += 1
                    continue

                # 4. Cor
                if valor_chave not in mapa_cores:
                    mapa_cores[valor_chave] = self.gerar_cor_pastel(DB)
                
                cor_usar = mapa_cores[valor_chave]

                # 5. Override (A MÁGICA ACONTECE AQUI)
                override = DB.OverrideGraphicSettings()
                
                # Define apenas o fundo (Background)
                override.SetSurfaceForegroundPatternId(padrao_solido.Id)
                override.SetSurfaceForegroundPatternColor(cor_usar)
                
                # CORREÇÃO: Removi a linha abaixo que pintava o texto
                # override.SetProjectionLineColor(cor_usar) 

                view_ativa.SetElementOverrides(tag.Id, override)
                count_modificados += 1

            t.Commit()
            
            # --- FEEDBACK ---
            msg = "Sucesso! {} tags coloridas.".format(count_modificados)
            if count_sem_param > 0:
                msg += "\n({} ignoradas/sem valor)".format(count_sem_param)
            
            self.janela.txtStatus.Text = msg
            self.janela.txtStatus.Foreground = self.janela.cor_sucesso

        except Exception as e:
            self.janela.txtStatus.Text = "Erro: " + str(e)
            self.janela.txtStatus.Foreground = self.janela.cor_erro
            print(str(e))
            traceback.print_exc()

    def GetName(self):
        return "ColorirTagsHandler"

    def gerar_cor_pastel(self, DB):
        import random
        from System import Byte
        
        r = random.randint(180, 255)
        g = random.randint(180, 255)
        b = random.randint(180, 255)
        return DB.Color(Byte(r), Byte(g), Byte(b))


# --- 2. A JANELA ---
class JanelaColorir(forms.WPFWindow):
    def __init__(self):
        try:
            forms.WPFWindow.__init__(self, 'script.xaml')
        except Exception as e:
            forms.alert("Erro ao carregar script.xaml:\n" + str(e))
            return
        
        from System.Windows.Media import Brushes
        self.cor_sucesso = Brushes.Green
        self.cor_erro = Brushes.Red
        
        self.handler = ColorirHandler(self)
        self.evento = ExternalEvent.Create(self.handler)

    def btn_colorir_click(self, sender, args):
        self.handler.modo = "APLICAR"
        self.handler.nome_parametro = self.inputParametro.Text
        self.evento.Raise()

    def btn_limpar_click(self, sender, args):
        self.handler.modo = "LIMPAR"
        self.evento.Raise()

# --- INICIAR ---
JanelaColorir().Show()