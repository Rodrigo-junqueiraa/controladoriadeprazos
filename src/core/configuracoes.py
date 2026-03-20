"""Configurações de prazos por ramo jurídico"""

PRAZOS = {
    "Direito Trabalhista - CLT": {
        "Agravo TST": 8,
        "Agravo de Instrumento em RR ou RO": 8,
        "Agravo de Petição": 8,
        "Contraminuta AI / Contrarrazões RR": 8,
        "Contrarrazões ao RR ou RO": 8,
        "Embargos TST": 8,
        "Embargos à Execução": 5,
        "Embargos de declaração": 5,
        "Impugnação à Sentença de Liquidação": 5,
        "Recurso de Revista": 8,
        "Recurso Ordinário Trabalhista": 8
    },
    "Direito Civil - CPC": {
        "Agravo de Instrumento": 15,
        "Apelação": 15,
        "Embargos de declaração": 5,
        "Recurso Especial": 15,
        "Recurso Extraordinário": 15,
        "Recurso Ordinário (Cível)": 15
    }
}

# Temas e estilos
TEMA = {
    "bg_principal": "#0f172a",
    "bg_secundario": "#1e293b",
    "fg_principal": "white",
    "cor_azul": "#2563eb",
    "cor_verde": "#22c55e",
    "cor_vermelho": "#dc2626",
    "cor_hover": "#3b82f6",
    "font_titulo": ("Segoe UI", 16, "bold"),
    "font_subtitulo": ("Segoe UI", 14, "bold"),
    "font_normal": ("Segoe UI", 11),
    "font_pequeno": ("Segoe UI", 9),
}

# Caminhos de arquivos
ARQUIVOS = {
    "advogados_json": "advogados.json",
    "prazos_json": "CAMINHO_JSON",
    "excel_modelo": "Planilha de prazos - atualizada.xlsx",
    "imagem": "justica.png",
    "icone_inicio": "icon_inicio.png",
    "icone_calc": "icon_calc.png",
    "icone_notif": "icon_notification.png",
    "icone_lista": "icon_listagem.png",
    "icone_config": "icon_config.png",
}
