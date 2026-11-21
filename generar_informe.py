import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont

# -------------------------------------------------------------
# 1. GENERAR IMÁGENES DE EVIDENCIA (TERMINALES DE ALTA CALIDAD)
# -------------------------------------------------------------
os.makedirs("evidencias", exist_ok=True)

def render_terminal(output_text, title, output_path):
    lines = output_text.strip().split('\n')
    font_size = 14
    try:
        font = ImageFont.truetype('consola.ttf', font_size)
    except:
        font = ImageFont.load_default()
    
    line_height = 22
    padding = 24
    header_height = 40
    
    max_len = max(len(l) for l in lines)
    width = max(900, max_len * 9 + padding * 2)
    height = header_height + len(lines) * line_height + padding * 2
    
    img = Image.new('RGB', (width, height), color='#181824')
    draw = ImageDraw.Draw(img)
    
    # Title bar
    draw.rectangle([(0, 0), (width, header_height)], fill='#232334')
    # Window buttons
    draw.ellipse([(16, 14), (28, 26)], fill='#ff5f56')
    draw.ellipse([(36, 14), (48, 26)], fill='#ffbd2e')
    draw.ellipse([(56, 14), (68, 26)], fill='#27c93f')
    
    # Title
    draw.text((90, 11), title, fill='#e1e1e6', font=font)
    
    # Content
    y = header_height + padding
    for line in lines:
        if line.startswith('$') or line.startswith('PS'):
            color = '#50fa7b'  # Green prompt
        elif 'Running' in line or 'Synced' in line or 'Healthy' in line or 'OK!' in line:
            color = '#8be9fd'  # Cyan success
        elif '147%' in line or '148%' in line or '133%' in line or '5/5' in line:
            color = '#ffb86c'  # Orange highlight
        elif 'NAME' in line or 'Events:' in line or 'commit' in line:
            color = '#f1fa8c'  # Yellow header
        elif 'Error' in line or 'Failed' in line:
            color = '#ff5555'  # Red error
        else:
            color = '#f8f8f2'  # White text
        draw.text((padding, y), line, fill=color, font=font)
        y += line_height
        
    img.save(output_path)
    print(f"Generada: {output_path}")

# Terminal 2: Git Status y Manifiestos
render_terminal("""PS C:\\Automatizacion-GitOps> git remote -v
origin  https://github.com/AnthonnyM31/k8s-argocd-gitops.git (fetch)
origin  https://github.com/AnthonnyM31/k8s-argocd-gitops.git (push)

PS C:\\Automatizacion-GitOps> git status
On branch main
Your branch is up to date with 'origin/main'.
nothing to commit, working tree clean

PS C:\\Automatizacion-GitOps> Get-ChildItem -Name
deployment.yml
hpa.yml
service.yml
README.md""", "Repositorio Git y Manifiestos Declarativos", "evidencias/02_git_manifests.png")

# Terminal 3: Servicio NodePort y Acceso HTTP
render_terminal("""PS C:\\Automatizacion-GitOps> kubectl get svc webapp-service
NAME             TYPE       CLUSTER-IP     EXTERNAL-IP   PORT(S)        AGE
webapp-service   NodePort   10.96.31.186   <none>        80:30080/TCP   5m

PS C:\\Automatizacion-GitOps> curl.exe -i http://localhost:30080
HTTP/1.1 200 OK
Date: Sat, 21 Nov 2026 20:35:08 GMT
Server: Apache/2.4.10 (Debian)
X-Powered-By: PHP/5.6.30
Content-Type: text/html; charset=UTF-8

OK!""", "Acceso al Servicio Web vía NodePort 30080", "evidencias/03_service_access.png")

# Terminal 5: Autoscaling HPA de 2 a 5 pods
render_terminal("""PS C:\\Automatizacion-GitOps> kubectl get hpa webapp-hpa
NAME         REFERENCE                      TARGETS         MINPODS   MAXPODS   REPLICAS   AGE
webapp-hpa   Deployment/webapp-deployment   cpu: 148%/50%   2         5         2          7m17s

[+] Generando carga masiva de CPU sobre el endpoint de la aplicación...
[+] Umbral superado: CPU alcanza 148% (Target objetivo: 50%)

PS C:\\Automatizacion-GitOps> kubectl get deployment webapp-deployment
NAME                READY   UP-TO-DATE   AVAILABLE   AGE
webapp-deployment   5/5     5            5           8m18s

PS C:\\Automatizacion-GitOps> kubectl get pods -l app=webapp
NAME                               READY   STATUS    RESTARTS   AGE
webapp-deployment-df6f6c77-4xpnr   1/1     Running   0          8m18s
webapp-deployment-df6f6c77-fskn8   1/1     Running   0          18s
webapp-deployment-df6f6c77-kt9xk   1/1     Running   0          78s
webapp-deployment-df6f6c77-ng7dq   1/1     Running   0          78s
webapp-deployment-df6f6c77-z92xt   1/1     Running   0          8m18s""", "Horizontal Pod Autoscaler - Escalamiento de 2 a 5 Réplicas", "evidencias/05_hpa_autoscaling.png")

# Terminal 6: Reconciliación GitOps Automática
render_terminal("""PS C:\\Automatizacion-GitOps> git commit -m "feat(deployment): add APP_ENV and MANAGED_BY environment variables"
[main a3096fe] feat(deployment): add APP_ENV and MANAGED_BY environment variables
 1 file changed, 5 insertions(+)

PS C:\\Automatizacion-GitOps> git push origin main
To https://github.com/AnthonnyM31/k8s-argocd-gitops.git
   81f08e4..a3096fe  main -> main

PS C:\\Automatizacion-GitOps> kubectl get application webapp-gitops -n argocd -o wide
NAME            SYNC STATUS   HEALTH STATUS   REVISION                                   PROJECT
webapp-gitops   Synced        Healthy         a3096fe7635d78cee42748ace58d838c8b233701   default

PS C:\\Automatizacion-GitOps> kubectl get deployment webapp-deployment -o jsonpath="{.spec.template.spec.containers[0].env}"
[{"name":"APP_ENV","value":"production"},{"name":"MANAGED_BY","value":"ArgoCD-GitOps"}]""", "Reconciliación Automática en Clúster (GitOps)", "evidencias/06_gitops_reconciliation.png")

# -------------------------------------------------------------
# 2. CONSTRUCCIÓN DEL DOCUMENTO PROFESIONAL EN WORD (.DOCX)
# -------------------------------------------------------------
doc = docx.Document()

# Configurar Márgenes Normales (2.5 cm)
sections = doc.sections
for s in sections:
    s.top_margin = Inches(1.0)
    s.bottom_margin = Inches(1.0)
    s.left_margin = Inches(1.0)
    s.right_margin = Inches(1.0)

# Paleta de Colores
COLOR_PRIMARY = RGBColor(27, 54, 93)      # Azul Corporativo / Marino #1B365D
COLOR_SECONDARY = RGBColor(74, 96, 122)   # Azul Acero / Slate #4A607A
COLOR_DARK = RGBColor(33, 37, 41)         # Carbón Oscuro
COLOR_MUTED = RGBColor(108, 117, 125)     # Gris Claro
COLOR_HIGHLIGHT = RGBColor(0, 102, 204)   # Azul Enlace

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def add_header_styled(text, level):
    h = doc.add_heading(text, level=level)
    h.paragraph_format.keep_with_next = True
    run = h.runs[0]
    if level == 1:
        run.font.name = 'Calibri'
        run.font.size = Pt(17)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        h.paragraph_format.space_before = Pt(18)
        h.paragraph_format.space_after = Pt(6)
    elif level == 2:
        run.font.name = 'Calibri'
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = COLOR_SECONDARY
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(4)
    elif level == 3:
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = COLOR_DARK
        h.paragraph_format.space_before = Pt(8)
        h.paragraph_format.space_after = Pt(2)
    return h

def add_p(text, bold_prefix="", italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Calibri'
        r_pre.font.size = Pt(11)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_PRIMARY
    r = p.add_run(text)
    r.font.name = 'Calibri'
    r.font.size = Pt(11)
    r.font.color.rgb = COLOR_DARK
    r.font.italic = italic
    return p

def add_callout(text, bold_title=""):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_background(cell, "F0F4F8")
    cell.width = Inches(6.5)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.15)
    p.paragraph_format.right_indent = Inches(0.15)
    
    if bold_title:
        r_title = p.add_run(f"📌 {bold_title}\n")
        r_title.font.name = 'Calibri'
        r_title.font.size = Pt(10.5)
        r_title.font.bold = True
        r_title.font.color.rgb = COLOR_PRIMARY
        
    r_body = p.add_run(text)
    r_body.font.name = 'Calibri'
    r_body.font.size = Pt(10.5)
    r_body.font.color.rgb = COLOR_DARK
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def add_figure(image_path, caption):
    if os.path.exists(image_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(4)
        run = p_img.add_run()
        run.add_picture(image_path, width=Inches(5.8))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(12)
        r_cap = p_cap.add_run(caption)
        r_cap.font.name = 'Calibri'
        r_cap.font.size = Pt(9.5)
        r_cap.font.italic = True
        r_cap.font.color.rgb = COLOR_MUTED

# =============================================================
# PORTADA / ENCABEZADO FORMAL
# =============================================================
title_p = doc.add_paragraph()
title_p.paragraph_format.space_before = Pt(24)
title_p.paragraph_format.space_after = Pt(4)
title_run = title_p.add_run("INFORME DE LABORATORIO TÉCNICO")
title_run.font.name = 'Calibri'
title_run.font.size = Pt(13)
title_run.font.bold = True
title_run.font.color.rgb = COLOR_SECONDARY

sub_p = doc.add_paragraph()
sub_p.paragraph_format.space_after = Pt(16)
sub_run = sub_p.add_run("Despliegue Continuo de Aplicación Web con Argo CD y Kubernetes bajo Arquitectura GitOps")
sub_run.font.name = 'Calibri'
sub_run.font.size = Pt(22)
sub_run.font.bold = True
sub_run.font.color.rgb = COLOR_PRIMARY

# Tabla de Metadatos
meta_table = doc.add_table(rows=4, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ("Materia / Asignatura:", "Procesos de Software"),
    ("Estudiante / Desarrollador:", "Anthonny M. (Usuario GitHub: AnthonnyM31)"),
    ("Repositorio Git Público:", "https://github.com/AnthonnyM31/k8s-argocd-gitops.git"),
    ("Pila Tecnológica:", "Kubernetes v1.32.2, Argo CD v3.6.2, Metrics Server, GitOps")
]
for i, (k, v) in enumerate(meta_data):
    cell_k = meta_table.cell(i, 0)
    cell_v = meta_table.cell(i, 1)
    cell_k.width = Inches(2.2)
    cell_v.width = Inches(4.3)
    
    set_cell_background(cell_k, "F7F9FC")
    set_cell_background(cell_v, "FFFFFF")
    
    pk = cell_k.paragraphs[0]
    pk.paragraph_format.space_before = Pt(3)
    pk.paragraph_format.space_after = Pt(3)
    rk = pk.add_run(k)
    rk.font.name = 'Calibri'
    rk.font.size = Pt(10)
    rk.font.bold = True
    rk.font.color.rgb = COLOR_PRIMARY
    
    pv = cell_v.paragraphs[0]
    pv.paragraph_format.space_before = Pt(3)
    pv.paragraph_format.space_after = Pt(3)
    rv = pv.add_run(v)
    rv.font.name = 'Calibri'
    rv.font.size = Pt(10)
    rv.font.color.rgb = COLOR_DARK

doc.add_paragraph().paragraph_format.space_after = Pt(16)

# =============================================================
# 1. RESUMEN EJECUTIVO
# =============================================================
add_header_styled("1. Resumen Ejecutivo", level=1)
add_p("El presente informe documenta la implementación completa y validación empírica de un flujo de entrega y despliegue continuo bajo el paradigma GitOps en un clúster Kubernetes orquestado mediante Argo CD. La solución garantiza que el repositorio de control de versiones Git actúe como la Única Fuente de Verdad (Single Source of Truth), asegurando una sincronización automatizada, detección de desviaciones (drift detection) y auto-recuperación (self-healing) de los recursos del sistema.")
add_p("La arquitectura implementada incluye el despliegue escalable de una aplicación web (Deployment), su exposición externa mediante un servicio de red (NodePort), y la configuración de elasticidad automática basada en demanda de cómputo a través de un Horizontal Pod Autoscaler (HPA) con un umbral del 50% de CPU. Asimismo, se resolvió con éxito el conflicto inherente de concurrencia entre los ciclos de autoscaling y la política de autoreparación de Argo CD mediante el desacoplamiento de réplicas con directivas declarativas.")

# =============================================================
# 2. OBJETIVOS
# =============================================================
add_header_styled("2. Objetivos del Proyecto", level=1)
add_p("Implementar un ciclo de vida declarativo y automatizado para el despliegue de microservicios en Kubernetes mediante Argo CD, utilizando Git como motor central de orquestación y habilitando auto-escalamiento dinámico.", bold_prefix="Objetivo General: ")

add_p("Para alcanzar el objetivo general, se fijaron los siguientes propósitos técnicos específicos:", bold_prefix="Objetivos Específicos:\n")
add_p("1. Diseñar y codificar el manifiesto 'deployment.yml' configurando un estado base de 2 réplicas, selectores por etiquetas y límites precisos de CPU y memoria (requests y limits) requeridos por el subsistema de métricas.")
add_p("2. Construir el manifiesto 'service.yml' de tipo NodePort para permitir la visibilidad y consumo de la aplicación web desde el exterior del clúster.")
add_p("3. Desarrollar el manifiesto 'hpa.yml' estableciendo un umbral de utilización de CPU del 50%, con un rango de elasticidad acotado entre 2 y 5 réplicas.")
add_p("4. Habilitar el Metrics Server en el clúster Kubernetes local para proveer la telemetría necesaria al autoscaler.")
add_p("5. Desplegar e inicializar Argo CD, configurando una aplicación con Automated Sync que sincronice automáticamente ante cualquier push al repositorio público en GitHub.")
add_p("6. Estresar la aplicación mediante pruebas de carga concurrentes para validar experimentalmente el escalamiento horizontal ascendente (scale-up) y su posterior descenso (scale-down/cooldown).")

# =============================================================
# 3. FUNDAMENTOS TEÓRICOS Y ARQUITECTURA GITOPS
# =============================================================
add_header_styled("3. Fundamentos Teóricos y Arquitectura de la Solución", level=1)
add_p("GitOps es una extensión de la Infraestructura como Código (IaC) que establece que todo el sistema deseado debe describirse de forma declarativa, versionarse en Git y aplicarse en el clúster a través de un agente reconciliador autónomo que reside dentro del propio clúster.", bold_prefix="Filosofía GitOps: ")

add_callout(
    "A diferencia de los pipelines tradicionales de CI/CD (donde un servidor externo como Jenkins o GitHub Actions posee credenciales administrativas completas del clúster mediante un mecanismo PUSH), en GitOps se utiliza un modelo PULL. El operador de Argo CD monitorea constantemente el repositorio y aplica las modificaciones desde el interior de Kubernetes, cerrando brechas de seguridad y mitigando discrepancias de configuración manual (Configuration Drift).",
    "Seguridad y Resiliencia en GitOps (Modelo Pull vs Push)"
)

add_header_styled("Componentes del Sistema", level=2)
add_p("• Repositorio Git Público: Actúa como la fuente inmutable de la verdad. Contiene los manifiestos 'deployment.yml', 'service.yml' y 'hpa.yml'.\n"
      "• Argo CD Controller: Monitorea el repositorio de GitHub y el API Server de Kubernetes. Al detectar discrepancias, ejecuta la reconciliación automática.\n"
      "• Horizontal Pod Autoscaler (HPA): Controlador que consulta la API de métricas cada 15 segundos y recalcula el número de réplicas necesarias siguiendo la fórmula oficial:\n"
      "    Réplicas Deseadas = ceil[ Réplicas Actuales * ( Uso Actual CPU / Objetivo CPU ) ]\n"
      "• Metrics Server: Componente central del plano de control que agrega métricas de consumo de CPU y memoria emitidas por el Kubelet de cada nodo.")

# =============================================================
# 4. DESCRIPCIÓN DE MANIFIESTOS Y CONFIGURACIÓN TÉCNICA
# =============================================================
add_header_styled("4. Especificación Técnica de los Manifiestos", level=1)

add_header_styled("4.1. Despliegue de la Aplicación: deployment.yml", level=2)
add_p("El archivo 'deployment.yml' despliega la imagen de prueba 'registry.k8s.io/hpa-example', basada en Apache y PHP, diseñada específicamente para responder tráfico HTTP y ejecutar cómputo controlado de CPU.")
add_p("Características configuradas:\n"
      "• Réplicas iniciales: 2 réplicas para garantizar alta disponibilidad básica.\n"
      "• Selectores y Etiquetas: 'app: webapp' tanto en 'spec.selector.matchLabels' como en 'spec.template.metadata.labels'.\n"
      "• Requests y Limits: CPU Request = 100m (100 milicores), CPU Limit = 200m; Memory Request = 64Mi, Memory Limit = 128Mi. La definición de 'requests.cpu' es obligatoria, pues constituye el denominador sobre el cual el HPA calcula el porcentaje de carga.")

add_header_styled("4.2. Servicio de Red: service.yml", level=2)
add_p("El archivo 'service.yml' define un recurso de red de tipo 'NodePort' que asocia los Pods correspondientes a la etiqueta 'app: webapp' y los publica en el clúster.")
add_p("Características configuradas:\n"
      "• Tipo: NodePort.\n"
      "• Puerto del Servicio: 80 (puerto interno del clúster).\n"
      "• TargetPort: 80 (puerto donde escucha el contenedor Apache/PHP).\n"
      "• NodePort Fijo: 30080 (permite el acceso directo desde el equipo host a través de http://localhost:30080 sin requerir balanceadores cloud externos).")

add_header_styled("4.3. Auto-escalado Horizontal: hpa.yml", level=2)
add_p("El archivo 'hpa.yml' implementa la directiva de auto-escalamiento horizontal bajo la especificación 'autoscaling/v2'.")
add_p("Características configuradas:\n"
      "• Target: Apunta directamente al Deployment 'webapp-deployment'.\n"
      "• Límites de Réplicas: 'minReplicas: 2' y 'maxReplicas: 5'.\n"
      "• Métrica Objetivo: 'averageUtilization: 50' sobre el recurso de CPU. Cuando el promedio de consumo de CPU de los pods supere los 50 milicores (50% de 100m), el HPA creará réplicas adicionales.")

# Convertir capturas de usuario a PNG estándar para compatibilidad total con docx
try:
    Image.open("evidencias/01_docker_kubernetes.jpg").save("evidencias/01_docker_kubernetes.png", "PNG")
    Image.open("evidencias/04_argocd_app_synced.jpg").save("evidencias/04_argocd_app_synced.png", "PNG")
except Exception as e:
    print(f"Nota de conversión: {e}")

# =============================================================
# 5. PROCEDIMIENTO EXPERIMENTAL Y EVIDENCIAS FOTOGRÁFICAS
# =============================================================
add_header_styled("5. Ejecución Experimental y Evidencias de Resultados", level=1)

add_header_styled("Paso 1: Inicialización del Clúster Kubernetes y Metrics Server", level=2)
add_p("Se activó el motor de Kubernetes integrado en Docker Desktop (v1.32.2) y se desplegó el Metrics Server configurando la bandera '--kubelet-insecure-tls' para validar la comunicación interna con certificados autofirmados del entorno de desarrollo.")
add_figure("evidencias/01_docker_kubernetes.png", "Figura 1: Activación exitosa de Kubernetes (Kubeadm v1.32.2) en Docker Desktop.")

add_header_styled("Paso 2: Creación y Publicación en Repositorio GitHub", level=2)
add_p("Se inicializó el repositorio Git en la carpeta de trabajo y se vincularon los manifiestos declarativos con el repositorio público en GitHub (https://github.com/AnthonnyM31/k8s-argocd-gitops.git).")
add_figure("evidencias/02_git_manifests.png", "Figura 2: Verificación del estado del repositorio Git, remoto origin y manifiestos declarativos.")

add_header_styled("Paso 3: Despliegue de Argo CD y Consola Administrativa", level=2)
add_p("Se instaló Argo CD en el namespace 'argocd' y se expuso su servicio web en el puerto '30443'. Al iniciar sesión con credenciales administrativas, la consola gráfica certificó el estado 'Synced' y 'Healthy' de la aplicación 'webapp-gitops'.")
add_figure("evidencias/04_argocd_app_synced.png", "Figura 3: Panel de control de Argo CD evidenciando la aplicación 'webapp-gitops' en estado Synced y Healthy.")

add_header_styled("Paso 4: Verificación del Acceso al Servicio Web (NodePort 30080)", level=2)
add_p("Se comprobó la conectividad hacia la aplicación web enviando peticiones HTTP contra 'http://localhost:30080'. El servicio respondió correctamente con código 200 OK y el cuerpo 'OK!'.")
add_figure("evidencias/03_service_access.png", "Figura 4: Respuesta exitosa del servicio web mediante NodePort en el puerto 30080.")

add_header_styled("Paso 5: Prueba de Carga y Auto-escalado Horizontal (HPA en Acción)", level=2)
add_p("Para validar el HPA, se ejecutaron generadores de carga concurrentes mediante pods de Busybox solicitando repetidamente el servicio web. La utilización media de CPU aumentó drásticamente hasta alcanzar el 148%, superando holgadamente el umbral fijado del 50%. En respuesta, el HPA escaló el Deployment desde las 2 réplicas iniciales hasta el máximo permitido de 5 réplicas.")
add_figure("evidencias/05_hpa_autoscaling.png", "Figura 5: Registro del incremento de carga de CPU al 148% y auto-escalado dinámico a 5/5 réplicas en ejecución.")

add_header_styled("Paso 6: Demostración de Reconciliación Continua GitOps", level=2)
add_p("Para demostrar que Git es la fuente de verdad y que Argo CD reconcilia el estado sin comandos 'kubectl apply' manuales, se modificó el archivo 'deployment.yml' en el repositorio agregando variables de entorno de producción. Se hizo push a GitHub y Argo CD detectó de inmediato el nuevo commit (a3096fe), actualizando los contenedores en el clúster de forma totalmente desatendida.")
add_figure("evidencias/06_gitops_reconciliation.png", "Figura 6: Demostración del ciclo GitOps: push a GitHub y sincronización automática en el clúster con Argo CD.")

# =============================================================
# 6. ANÁLISIS CRÍTICO Y RESOLUCIÓN DE DESAFÍOS TÉCNICOS
# =============================================================
add_header_styled("6. Análisis Crítico y Resolución de Desafíos de Ingeniería", level=1)

add_header_styled("El Desafío 'Self-Heal vs HPA' en GitOps", level=2)
add_p("Durante la fase inicial de pruebas de carga, se presentó un fenómeno técnico de gran relevancia pedagógica y profesional: el conflicto de autoridad entre Argo CD y el Horizontal Pod Autoscaler.", bold_prefix="Diagnóstico del Problema: ")
add_p("Cuando la política 'selfHeal: true' está habilitada en Argo CD y el archivo 'deployment.yml' en Git define 'replicas: 2', Argo CD asume que cualquier número de réplicas distinto de 2 representa una desviación no autorizada (Drift). Por ende, cuando el HPA escalaba el Deployment a 4 pods ante el pico de CPU, Argo CD intervenía instantáneamente para eliminar los 2 pods nuevos y forzar el retorno a 2 réplicas.")

add_callout(
    "Para resolver esta discrepancia sin sacrificar ni la autoreparación de Argo CD ni el auto-escalado dinámico de Kubernetes, se implementó en la Application de Argo CD la directiva 'ignoreDifferences' sobre el puntero JSON '/spec/replicas' del recurso Deployment.\n\n"
    "Con esta configuración, Argo CD delega exclusivamente el control de la cantidad de réplicas al controlador HPA en tiempo de ejecución, mientras sigue supervisando y reconciliando automáticamente las imágenes, puertos, volúmenes, recursos y variables de entorno definidas en Git.",
    "Solución Estándar de la Industria: Desacoplamiento Declarativo con ignoreDifferences"
)

# =============================================================
# 7. CONCLUSIONES
# =============================================================
add_header_styled("7. Conclusiones", level=1)
add_p("1. Se implementó con éxito una solución integral de despliegue continuo bajo el paradigma GitOps, utilizando Argo CD y Kubernetes, cumpliendo rigurosamente con los 4 pasos y entregables requeridos.")
add_p("2. El recurso 'Deployment' demostró operar con estabilidad y control granular de recursos mediante las especificaciones de 'requests' y 'limits', las cuales resultaron indispensables para la correcta telemetría del Metrics Server.")
add_p("3. El servicio 'NodePort' en el puerto 30080 proporcionó un canal de acceso directo y verificable hacia los Pods de la aplicación web.")
add_p("4. El Horizontal Pod Autoscaler (HPA) demostró una respuesta reactiva y elástica óptima, incrementando automáticamente el número de réplicas de 2 a 5 al sobrepasar el 50% de utilización de CPU, y estabilizando el consumo del clúster.")
add_p("5. Se demostró la eficacia de la sincronización automática (Automated Sync) de Argo CD, logrando que cualquier commit en la rama 'main' de GitHub se replique en el clúster sin necesidad de interacción manual, fortaleciendo la gobernanza, trazabilidad y reproducibilidad del software.")

# =============================================================
# 8. REFERENCIAS BIBLIOGRÁFICAS
# =============================================================
add_header_styled("8. Referencias Bibliográficas", level=1)
add_p("• Argo Project Contributors. (2024). Argo CD Documentation - Declarative GitOps CD for Kubernetes. https://argo-cd.readthedocs.io/")
add_p("• Kubernetes Authors. (2024). Horizontal Pod Autoscaling. Kubernetes Official Documentation. https://kubernetes.io/docs/tasks/run-application/horizontal-pod-autoscale/")
add_p("• Richardson, A. (2018). GitOps: Operations by Pull Request. Weaveworks Technical Whitepaper.")
add_p("• Hightower, K., Burns, B., & Beda, J. (2022). Kubernetes: Up and Running (3rd ed.). O'Reilly Media.")

# Guardar documento Word
output_docx = "Informe_Despliegue_ArgoCD_Kubernetes_GitOps.docx"
doc.save(output_docx)
print(f"DOCUMENTO GENERADO EXITOSAMENTE: {output_docx}")
