# Librerias e insumos
import streamlit as st
import pandas as pd
from datetime import datetime, date
import uuid # para crear ids únicas
from PIL import Image
from shareplum import Office365
from shareplum import Site
from shareplum.site import Version
from io import StringIO
import pytz
image = Image.open('logo mutual.jpg') # Guardar imagen del logo mutual en la misma carpeta de la App

# Código para abrir app en la ventana de comandos de Anaconda
# $ cd C:\Users\lalarcon\OneDrive - Mutual\Escritorio\Luis\Proyectos Data Science\EP-ME\App (cambiar por directorio personal)
# $ streamlit run Epicondilitis.py

diagnostico='Bursitis Subacromial de Hombro' # Modificar diagnóstico

# Titulo
st.title('Calificador Automático de EP-ME')

# Subtitulo
c1,c2,c3=st.columns(3)
with c2:
    st.write(" ###### Mutual de Seguridad CChC")

# 1. Selección de Microlabor o Macrolabor
labor = st.sidebar.selectbox(
    '¿La tarea corresponde a una Macrolabor o Microlabor?',
    ('Macrolabor', 'Microlabor'))

# 2. Tiempo total de la jornada (TTJ) 
fecha_informe = st.date_input('Fecha Informe', date.today())
if labor=='Macrolabor': # Solo aplica para Macrolabor
    tiempo_total = st.number_input(
        '¿Cuánto es el **tiempo total** -en minutos- de la jornada de trabajo (TTJ)?',
        min_value=0,  step=10)
else:
    tiempo_total='No aplica'

# 3.1 Selección del número de tareas -> Permite 5 tareas
n_tareas = st.sidebar.radio(
    '¿Cuántas tareas realiza el trabajador?',
    (1, 2, 3, 4, 5)) # Se definió un máximo de 5 tareas, pero podrían llegar a ser más

# 3.2 Datos paciente -> Para futuros cruces con otras bases de datos
rut_paciente = st.sidebar.text_input('Rut del trabajador (Sin puntos y con guión)','12345678-9')
rut_empresa = st.sidebar.text_input('Rut del empleador (Sin puntos y con guión)','12345678-9')

# 4. Preguntas
tabla=[] # Lista para pegar los datos en un dataframe
Td, Tm=[], [] # Listado de tiempos dedicados a cada tipo de tarea [diarias (Td) / no diarias (Tm)]
listado_tareas_con_riesgo_td, listado_tareas_con_riesgo_tm=[], [] # Listas para almacenar número de tareas diarias y no diarias
for i in range(n_tareas): # Se activa for/loop por cada tarea
    resultado=0 # Indicador del puntaje obtenido
    st.write("### Respecto a la tarea",i+1,'responda:') # Enumeración de cada tarea
    
    # 4.1 Tiempo de trabajo diario (Td) o mensual (Tm) con exposición a riesgo
    if labor=='Macrolabor':
        variabilidad_tarea= st.selectbox('¿La tarea '+str(i+1)+' es realizada **todos los días de la semana**?',
            ['Si', 'No'])
        if variabilidad_tarea=='No':
            tiempo_tarea_mensual= st.number_input('¿Cúanto es el **tiempo mensual** -en minutos- de trabajo con exposición a riesgo (Tm) dedicado a la tarea '+str(i+1)+'?',
            min_value=0.0,  step=10.0, max_value=float(10800))
            Tm.append(tiempo_tarea_mensual)

            tiempo_tarea_semanal='No aplica'
        else:
            tiempo_tarea_semanal = st.number_input('¿Cuánto es el **tiempo diario** -en minutos- de trabajo con exposición a riesgo (Td) dedicado a la tarea '+str(i+1)+'?',
            min_value=0.0,  step=10.0, max_value=float(tiempo_total))
            Td.append(tiempo_tarea_semanal)

            tiempo_tarea_mensual='No aplica'
    else:
        variabilidad_tarea='Si' # Formula diferenciadora de tareas diarias y no diarias solo aplica para Macrolabores. Para simplicidad del código, se asumió que la Microlabor solo tiene tareas diarias
        tiempo_tarea_semanal = st.number_input('¿Cuánto es el **tiempo diario** -en minutos- de de trabajo con exposición a riesgo (Td) dedicado a la tarea '+str(i+1)+'?',
            min_value=0.0,  step=10.0)
        Td.append(tiempo_tarea_semanal)

        tiempo_tarea_mensual='No aplica'

    # 4.2 Abducción del hombro
    st.markdown("### Pregunta 1")
    abd_hombro = st.slider(
        '¿Cuál es el rango de amplitud -en grados- de **abducción** de hombro al realizar la tarea '+str(i+1)+'?',
        0, 100, (0, 59))
    'Abducción máxima:', abd_hombro[1] # Mostrar máxima abducción de hombro. Éste será el valor tomado para calcular el riesgo

    if abd_hombro[1]<60:
        resultado=resultado+0 # Si la abducción es menor a 60°, el puntaje será 0
    elif (abd_hombro[1]>=60) & (abd_hombro[1]<=89):
        resultado=resultado+1 # Si la abducción esta en el intervalo [60°,89°], el puntaje será 1
    elif abd_hombro[1]>=90:
        resultado=resultado+2 # Si abducción es mayor o igual a 90°, el puntaje será 2

    # 4.3 Rotación interna/externa
    st.markdown("### Pregunta 2")
    rotacion = st.selectbox(
        '¿Se observa **rotación interna y/o externa** del hombro al realizar la tarea '+str(i+1)+'?',
        ['Ausente', 'Presente'])

    if rotacion=='Ausente':
        resultado=resultado+0 # Si la rotación no esta presente, el puntaje será 0
    elif (rotacion=='Presente') & (abd_hombro[1]==0):
        resultado=resultado+0
        mensaje='Se observa rotación, pero sin abducción de hombro. Por lo tanto, el puntaje será 0:'
        txt='<p style="font-family:Courier; color: Red; font-size: 12px">'+mensaje+'</p>'
        st.markdown(txt, unsafe_allow_html=True)
    elif (rotacion=='Presente') & (abd_hombro[1]>0): # Debe ir acompañada de abducción de hombro, independiente de su amplitud
        resultado=resultado+1 # Si la rotación esta presente y existe abducción, el puntaje será 1

    # 4.4 Postura mantenida
    st.markdown("### Pregunta 3")
    postura = st.selectbox(
        '¿Se observa **postura mantenida** mayor a 4 segundos -en abducción y/o rotación interna de hombro- al realizar la tarea '+str(i+1)+'?',
        ['Ausente', 'Presente'])

    if postura=='Ausente':
        resultado=resultado+0 # Si la postura no se mantiene por más de 4 segundos, el puntaje será 0
    elif postura=='Presente':
        resultado=resultado+1 # Si la postura se mantiene más de 4 segundos, el puntaje será 1

    # 4.4 Repetitividad
    st.markdown("### Pregunta 4")
    if labor=="Microlabor": # Esta pregunta solo aplica para una Microlabor
        repetitividad_mic = st.number_input(
        '¿Cuánto es el número **más alto** de repeticiones (Mov/Min) realizadas por el trabajador en la tarea '+str(i+1)+'?',
        min_value=0,  step=1)

        if repetitividad_mic<1:
            resultado=resultado+0 # Si la repetitividad de movimientos por minúto es menor a 1, el puntaje será 0
        elif (repetitividad_mic>=1) & (repetitividad_mic<3):
            resultado=resultado+1 # Si la repetitividad de movimientos por minúto se encuentra en el intervalo [1,3[, el puntaje será 1
        elif (repetitividad_mic>=3) & (repetitividad_mic<5):
            resultado=resultado+2 # Si la repetitividad de movimientos por minúto se encuentra en el intervalo [3,5[, el puntaje será 1
        elif repetitividad_mic>=5:
            resultado=resultado+3 # Si la repetitividad de movimientos por minúto es mayor o igual a 5, el puntaje será 2

        repetitividad_mac='No aplica' # Este item no aplica para macrolabor

    else:
        repetitividad_mac = st.selectbox(
            '¿Se observa repetitividad en la tarea '+str(i+1)+'?',
            ['Ausente', 'Presente'])

        if repetitividad_mac=='Ausente':
            resultado=resultado+0 # Si no se observa repetitividad, el puntaje será 0
        elif repetitividad_mac=='Presente':
            resultado=resultado+1 # Si se observa repetitividad,, el puntaje será 1

        repetitividad_mic='No aplica' # Este item no aplica para microlabor

    # 4.5 Fuerza (borg)
    st.markdown("### Pregunta 5")
    borg = st.slider(
    '¿Cuál es la **más alta** percepción de fuerza (Borg) realizada por el trabajador en la tarea '+str(i+1)+'?',
    0, 10, step=1)

    if borg<3:
        resultado=resultado+0 # Si el Borg es menor a 3, el puntaje será 0
    elif (borg>=3) & (borg<=4):
        resultado=resultado+1 # Si el Borg se encuentra en el intervalo [3,4], el puntaje será 1
    elif borg>4:
        resultado=resultado+2 # Si el Borg es mayor a 4, el puntaje será 2

    '#### El puntaje de la tarea '+str(i+1)+' es ', resultado # Variable resultado almacena puntaje final de cada tarea

    nivel_riesgo='Sin Riesgo'
    if resultado==0: 
        nivel_riesgo='Sin Riesgo'
        color='Grey'
        if (variabilidad_tarea=='Si'): # Para tareas diarias se guarda codificación de resultados en la lista de tareas diarias. Este proceso se repite en los siguientes resultados
            listado_tareas_con_riesgo_td.append(0) # Código 0: Sin riesgo o Insuficiente
        elif variabilidad_tarea=='No': # Por su parte, los resultados de las tareas no diarias se almacenan en la lista de tareas no diarias
            listado_tareas_con_riesgo_tm.append(0)
    elif ((resultado>=1) & (resultado<=2)): 
        nivel_riesgo='Insuficiente por puntaje'
        color='Grey'
        if (variabilidad_tarea=='Si'):
            listado_tareas_con_riesgo_td.append(0) # Código 0: Sin riesgo o Insuficiente
        elif variabilidad_tarea=='No':
            listado_tareas_con_riesgo_tm.append(0)
    elif resultado==3:
        nivel_riesgo='Leve'
        color='Green'
        if (variabilidad_tarea=='Si'):
            listado_tareas_con_riesgo_td.append(1) # Código 1: Riesgo Leve
        elif variabilidad_tarea=='No':
            listado_tareas_con_riesgo_tm.append(1)
    elif resultado==4:
        nivel_riesgo='Moderado'
        color='Orange'
        if (variabilidad_tarea=='Si'):
            listado_tareas_con_riesgo_td.append(2) # Código 2: Riesgo Moderado
        elif variabilidad_tarea=='No':
            listado_tareas_con_riesgo_tm.append(2)
    elif resultado>4:
        nivel_riesgo='Severo'
        color='Red'
        if (variabilidad_tarea=='Si'):
            listado_tareas_con_riesgo_td.append(3) # Código 3: Riesgo Severo
        elif variabilidad_tarea=='No':
            listado_tareas_con_riesgo_tm.append(3)

    texto_nivel_riesgo= '<p style="font-family:Courier; color:'+color+'; font-size: 20px">'+nivel_riesgo+'</p>'
    'Por lo tanto, su nivel de riesgo es : '
    st.markdown(texto_nivel_riesgo, unsafe_allow_html=True)

    data = {'labor': labor, 
            'n_tareas': n_tareas,
            'correlativo_tarea':i+1,
            'ttj':tiempo_total,
            'tarea_semanal':variabilidad_tarea,
            'tiempo_tarea_mensual':tiempo_tarea_mensual,
            'tiempo_tarea_semanal':tiempo_tarea_semanal,
            'extension_muneca':'No aplica',
            'flexion_muneca':'No aplica',
            'flexion_hombro':'No aplica',
            'golpe_mano':'No aplica',
            'pinzamiento':'No aplica',
            'abduccion_hombro':abd_hombro[1],
            'rotacion':rotacion,
            'supinacion':'No aplica',
            'pronacion':'No aplica',
            'postura':postura,
            'repetitividad_mic':repetitividad_mic,
            'repetitividad_mac':repetitividad_mac,
            'latko':'No aplica',
            'borg':borg,
            'puntaje':resultado,
            'riesgo':nivel_riesgo} 

    tabla.append(pd.DataFrame(data, index=[0])) 
    '---'

df=pd.concat(tabla)
df.insert(0,'id_registro',str(uuid.uuid4().fields[-1])[:5])
df.reset_index(inplace=True,drop=True)
# 5. Cálculo del RMac y RMic

# --Funciones--
def find_indices(item_to_find,tiempo): # Función para identificar riesgo de cada tarea de acuerdo a su codificación
    indices = []
    if tiempo=='Td':
        for idx, value in enumerate(listado_tareas_con_riesgo_td):
            if value == item_to_find: # Acá se identifica el código de cada tarea
                indices.append(idx) # Las tareas con riesgos equivalentes son almacenadas en listas de un mismo tipo (i.e. tareas severas se juntan en una lista de tareas severas y así sucesivamente)
        return indices
    elif tiempo=='Tm': # Se hace diferenciación entre tareas diarias y no diarias. De modo que tareas severas diarias van en una lista de tareas severas diarias y las no diarias tienen sus propias listas para cada riesgo
        for idx, value in enumerate(listado_tareas_con_riesgo_tm):
            if value == item_to_find:
                indices.append(idx)
        return indices

def calculo_rmac(posicion,rango_minimo,rango_intermedio,tiempo):
    T=[] # En esta lista se almacenan los tiempos dedicados a cada tarea
    for i in range(len(posicion)): # La posición permite calcular RMac para tareas con riesgos equivalente (e.g. tareas con riesgo moderado)
        if tiempo=='Td': # Se hace diferenciación por tarea diaria
            try:
                T.append(Td[posicion[i]]) # Se añaden los tiempo de cada tarea a la lista pertinente
            except:
                continue # Si la formula previa da error (e.g. no hay tareas con riesgo severo), entonces el loop se detiene 
        elif tiempo=='Tm': # Acá la lógica es la misma, pero para tareas no diarias
            try:
                T.append(Tm[posicion[i]])
            except:
                continue
    if tiempo=='Td':
        try:
            rmac=round((sum(T)/tiempo_total)*100,1) # Si la tarea es diaria se aplica la fórmula (Td/TTj)*100
        except:
            rmac=0 # Si fórmula anterior da error (e.g. denominador es 0), se asume que el RMac es 0. Esto ocurre cuando no se contesta la pregunta 2
    elif tiempo=='Tm':
        ttjm=10800 # Este valor es fijo y corresponde al número de minuto trabajados durante el mes para una persona que trabaja 45 horas semanales (45 horas * 60 minutos * 4 semanas = 10.800 minutos)
        rmac=round((sum(T)/ttjm)*100,1) # Si la tarea no es diaria se aplica la fórmula (Tm/TTjm)*100
    
    'El RMac es: ', rmac,'%, por lo que el caso califica como: '
    if rmac<rango_minimo: # Acá se define intervalo para decidir si enfermedad es de origen comun o profesional. Este intervalo va a cambiar dependiendo del tipo de riesgo. Es por ello que se piden los intervalos en el input de la función
        color='Green' # Se despliega texto color verde para enfermedades de origen comun
        resultado='Enfermedad comun'
        texto='<p style="font-family:Courier; color:'+color+'; font-size: 20px">'+resultado+'</p>' # Editar fuente de texto
        return st.markdown(texto, unsafe_allow_html=True),resultado,rmac # Desplegar texto
    elif (rmac>=rango_minimo) & (rmac<rango_intermedio): # Rango para situación límite
        color='Orange' # Se despliega texto color naranjo para situaciones límite
        resultado='Situación Límite'
        texto='<p style="font-family:Courier; color:'+color+'; font-size: 20px">'+resultado+'</p>' # Editar fuente de texto
        return st.markdown(texto, unsafe_allow_html=True),resultado,rmac # Desplegar texto
    else: # Rango para enfermedad profesional
        color='Red' # Se despliega texto color rojo para enfermedades profesionales
        resultado='Enfermedad Profesional'
        texto='<p style="font-family:Courier; color:'+color+'; font-size: 20px">'+resultado+'</p>' # Editar fuente de texto
        return st.markdown(texto, unsafe_allow_html=True),resultado,rmac # Desplegar texto

def calculo_rmic(posicion,rango_minimo,rango_intermedio): # Esta fórmula es similar a la anterior pero más simple, pues no requiere diferenciar por tareas diaras y no diarias
    T_=[]
    for i in range(len(posicion)):
        T_.append(Td[posicion[i]])
    rmic=round((sum(T_)/60),1) # Acá la fórmula es la suma de horas dedicadas a cada tareas llevada a minutos (por eso se divide por 60)
    'El RMic es: ', rmic,'horas, por lo que el caso califica como: '
    if rmic<rango_minimo:
        color='Green'
        resultado='Enfermedad comun'
        texto='<p style="font-family:Courier; color:'+color+'; font-size: 20px">'+resultado+'</p>'
        return st.markdown(texto, unsafe_allow_html=True),resultado,rmic
    elif (rmic>=rango_minimo) & (rmic<rango_intermedio):
        color='Orange'
        resultado='Situación Límite'
        texto='<p style="font-family:Courier; color:'+color+'; font-size: 20px">'+resultado+'</p>'
        return st.markdown(texto, unsafe_allow_html=True),resultado,rmic
    else:
        color='Red'
        resultado='Enfermedad Profesional'
        texto='<p style="font-family:Courier; color:'+color+'; font-size: 20px">'+resultado+'</p>'
        return st.markdown(texto, unsafe_allow_html=True),resultado,rmic

def resultado_rmac(tiempo): # Si el input (tiempo) es "Td" (i.e. tarea es diaria), entonces se ejecutará fórmula de RMac para tareas diarias. En caso contrario (e.g. tiempo = "Tm"), se ejecutará la fórmula RMac la tareas no diarias
    df['rmic__tareas_severas'], df['rmic__tareas_moderadas'], df['rmic__tareas_leves']='No aplica', 'No aplica', 'No aplica'    
    codigo_riesgo_tarea = 3 # Codificación 3: Riesgo Severo
    posicion=find_indices(codigo_riesgo_tarea,tiempo) # Se buscan tareas con riesgo severo
    if len(posicion)>0: # Si se detectan tareas con riesgo severo, entonces se cálcula RMac para dichas tareas
    
        st.markdown('### Cálculo del RMac para tareas con riesgo severo') # Desplegar texto
        res,calificacion,rmac=calculo_rmac(posicion,25,30,tiempo) # Se activa función para calcular RMac Severo
        df['rmac__tareas_severas'], df['calif__tareas_severas']= rmac, calificacion 
        if calificacion=='Enfermedad comun': # Si califica como enfermedad comun, entonces se procede a calcular RMac de tareas con riesgo moderado y asi sucesivamnte
            codigo_riesgo_tarea = 2 # Codificación 2: Riesgo Moderado
            posicion=find_indices(codigo_riesgo_tarea,tiempo) # Se buscan tareas con riesgo moderado
            if len(posicion)==0: # Si no se detectan tareas con riesgo moderado, se procede a identificar tareas con riesgo leve
                df['rmac__tareas_moderadas'], df['calif__tareas_moderadas']='No aplica', 'No aplica' 
                codigo_riesgo_tarea = 1 # Codificación 1: Riesgo Leve
                posicion=find_indices(codigo_riesgo_tarea,tiempo) # Se buscan tareas con riesgo leve
                if len(posicion)==0: # Si no hay tareas de riesgo leve, entonces no hay más tareas a evaluar y se acaba el proceso
                    df['rmac__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
                    'No hay más tareas a evaluar'
                elif len(posicion)>0: # Si se detectan tareas con riesgo leve, entonces se cálcula RMac para dichas tareas
                    st.markdown('### Cálculo del RMac para tareas con riesgo leve') # Desplegar texto
                    res,calificacion,rmac=calculo_rmac(posicion,65,70,tiempo) # Se activa función para calcular RMac Leve
                    df['rmac__tareas_leves'], df['calif__tareas_leves']= rmac, calificacion
                    if calificacion=='Enfermedad comun': # Si todos los RMac califican como enfermedad comun, entonces se acaba el proceso
                        'No hay más tareas a evaluar'
            elif len(posicion)>0: # Si se detectan tareas con riesgo moderado, entonces se cálcula RMac para dichas tareas
                st.markdown('### Cálculo del RMac para tareas con riesgo moderado') # Desplegar texto
                res,calificacion,rmac=calculo_rmac(posicion,45,50,tiempo) # Se activa función para calcular RMac Moderado
                df['rmac__tareas_moderadas'], df['calif__tareas_moderadas']= rmac, calificacion
                if calificacion=='Enfermedad comun': # Si califica como enfermedad comun, entonces se procede a calcular RMac de tareas con riesgo leve
                    codigo_riesgo_tarea = 1 # Codificación 1: Riesgo Leve
                    posicion=find_indices(codigo_riesgo_tarea,tiempo) # Se buscan tareas con riesgo leve
                    if len(posicion)==0: # Si no hay tareas de riesgo leve, entonces no hay más tareas a evaluar y se acaba el proceso
                        df['rmac__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
                        'No hay más tareas a evaluar'
                    elif len(posicion)>0: # Si se detectan tareas con riesgo leve, entonces se cálcula RMac para dichas tareas
                        st.markdown('### Cálculo del RMac para tareas con riesgo leve') # Desplegar texto
                        res,calificacion,rmac=calculo_rmac(posicion,65,70,tiempo)
                        df['rmac__tareas_leves'], df['calif__tareas_leves']= rmac, calificacion
                        if calificacion=='Enfermedad comun': # Si no hay tareas de riesgo leve, entonces no hay más tareas a evaluar y se acaba el proceso
                            'No hay más tareas a evaluar'
                else:
                    'No hay más tareas a evaluar'
                    df['rmac__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
        else:
            'No hay más tareas a evaluar'
            df['rmac__tareas_moderadas'], df['calif__tareas_moderadas']='No aplica', 'No aplica' 
            df['rmac__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'

    else: # Si no se detectan tareas con riesgo severo, se partirán buscando tareas con riesgo moderado y se repetirá el algoritmo de la sección anterior
        df['rmac__tareas_severas'], df['calif__tareas_severas']='No aplica', 'No aplica'
        codigo_riesgo_tarea = 2 
        posicion=find_indices(codigo_riesgo_tarea,tiempo)
        if len(posicion)>0:
            st.markdown('### Cálculo del RMac para tareas con riesgo moderado')
            res,calificacion,rmac=calculo_rmac(posicion,45,50,tiempo)
            df['rmac__tareas_moderadas'], df['calif__tareas_moderadas']= rmac, calificacion
            if calificacion=='Enfermedad comun':
                codigo_riesgo_tarea = 1
                posicion=find_indices(codigo_riesgo_tarea,tiempo)
                if len(posicion)>0:
                    st.markdown('### Cálculo del RMac para tareas con riesgo leve')
                    res,calificacion,rmac=calculo_rmac(posicion,65,70,tiempo)
                    df['rmac__tareas_leves'], df['calif__tareas_leves']= rmac, calificacion
                    if calificacion=='Enfermedad comun':
                        'No hay más tareas a evaluar'
                else:
                    df['rmac__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
                    'No hay más tareas a evaluar'
            else:
                df['rmac__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
                'No hay más tareas a evaluar'
        else:
            df['rmac__tareas_moderadas'], df['calif__tareas_moderadas']='No aplica', 'No aplica' 
            codigo_riesgo_tarea = 1 
            posicion=find_indices(codigo_riesgo_tarea,tiempo)
            if len(posicion)>0:
                st.markdown('### Cálculo del RMac para tareas con riesgo leve')
                res,calificacion,rmac=calculo_rmac(posicion,65,70,tiempo)
                df['rmac__tareas_leves'], df['calif__tareas_leves']= rmac, calificacion
                if calificacion=='Enfermedad comun':
                    'No hay más tareas a evaluar'
                else:
                    'No hay más tareas a evaluar'
            else:
                df['rmac__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
                'No hay tareas a evaluar'

# ----------------------------------------------------------------------------------------------------------

# Algoritmo para calcular RMac y RMic de acuerdo al árbol de decisión de la circular SUSESO
if labor=='Macrolabor':
    df['rmic__tareas_severas'], df['rmic__tareas_moderadas'], df['rmic__tareas_leves']='No aplica', 'No aplica', 'No aplica' 
    if (sum(listado_tareas_con_riesgo_td)==0) & (sum(listado_tareas_con_riesgo_tm)==0): # Si no se registran tareas con riesgo (codificación es 0 para cada tarea), entonces no se ejecuta fórmula de RMac
        'No hay tareas con riesgo'
        df['rmac__tareas_severas'], df['calif__tareas_severas']='No aplica', 'No aplica'
        df['rmac__tareas_moderadas'], df['calif__tareas_moderadas']='No aplica', 'No aplica' 
        df['rmac__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'

    if (sum(Td)==0) & (sum(Tm)==0): # Si no se registra tiempo destinada a cada tarea, no se ejecuta fórmula de RMac
        'No se ha registrado el tiempo destinado a cada tarea'
        df['rmac__tareas_severas'], df['calif__tareas_severas']='No aplica', 'No aplica'
        df['rmac__tareas_moderadas'], df['calif__tareas_moderadas']='No aplica', 'No aplica' 
        df['rmac__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'

    elif (len(Tm)==0) & (len(Td)>0) & (sum(listado_tareas_con_riesgo_td)>0): # Si hay tareas con riesgo y todas son diarias, se ejecutará solo RMac para tareas diarias
        color='Purple'
        mensaje='Todas las tareas se realizan diariamente. Por lo tanto, el RMac se calculó con la fórmula (Td/TTj)*100'
        txt='<p style="font-family:Courier; color:'+color+'; font-size: 20px">'+mensaje+'</p>'
        st.markdown(txt, unsafe_allow_html=True)
        resultado_rmac('Td')

    elif (len(Tm)>0) & (len(Td)==0) & (sum(listado_tareas_con_riesgo_tm)>0):  # Si hay tareas con riesgo y ninguna es diaria, se ejecutará solo RMac para tareas no diarias
        color='Purple'
        mensaje='Ninguna de las tareas se realizan diariamente. Por lo tanto, el RMac se calculó con la fórmula (Tm/TTjm)*100'
        txt='<p style="font-family:Courier; color:'+color+'; font-size: 20px">'+mensaje+'</p>'
        st.markdown(txt, unsafe_allow_html=True)
        resultado_rmac('Tm')

    elif (len(Tm)>0) & (len(Td)>0) & ((sum(listado_tareas_con_riesgo_td)>0) | (sum(listado_tareas_con_riesgo_tm)>0)): # Si hay tareas con riesgos y pueden ser diarias o no diarias, se ejecutará la fórumla RMac correspondiente a cada tipo de tarea
        color='Purple'
        mensaje='Hay tareas que se realizan diariamente y otras que no. Por lo tanto, el RMac se calculó con la fórmula (Td/TTj)*100 para tareas diarias y (Tm/TTjm)*100 para no diarias'
        txt='<p style="font-family:Courier; color:'+color+'; font-size: 20px">'+mensaje+'</p>'
        st.markdown(txt, unsafe_allow_html=True)

        st.markdown('## Tareas diarias') # Primero de calcula el RMac de las tareas diarias
        resultado_rmac('Td')
        
        st.markdown('## Tareas no diarias') # Despues el RMac de las tareas no diarias
        resultado_rmac('Tm')
     
else: # El loop de la Microlabor es equivalente al de Macrolabor, solo que más sencillo pues no requiere diferenciar por tareas diarias y no diarias
    df['rmac__tareas_severas'], df['rmac__tareas_moderadas'], df['rmac__tareas_leves']='No aplica', 'No aplica', 'No aplica' 
    codigo_riesgo_tarea = 3
    posicion=find_indices(codigo_riesgo_tarea,'Td')
    if len(posicion)==0:
        df['rmic__tareas_severas'], df['calif__tareas_severas']='No aplica', 'No aplica'
        codigo_riesgo_tarea = 2
        posicion=find_indices(codigo_riesgo_tarea,'Td')
        if len(posicion)==0:
            df['rmic__tareas_moderadas'], df['calif__tareas_moderadas']='No aplica', 'No aplica' 
            codigo_riesgo_tarea = 1
            posicion=find_indices(codigo_riesgo_tarea,'Td')
            if len(posicion)==0:
                df['rmic__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
                'No hay tareas con riesgos'
            elif len(posicion)>0:
                st.markdown('### Cálculo del RMic para tareas con riesgo leve')
                res,calificacion,rmic=calculo_rmic(posicion,3.5,4) # Rango minimo: 3.5 ; Rango intermedio: 4
                df['rmic__tareas_leves'], df['calif__tareas_leves']= rmic, calificacion
                if calificacion=='Enfermedad comun':
                    'No hay más tareas a evaluar' 
                
        elif len(posicion)>0:
            st.markdown('### Cálculo del RMic para tareas con riesgo moderado')
            res,calificacion,rmic=calculo_rmic(posicion,2.5,3) # Rango minimo: 2.5 ; Rango intermedio: 3
            df['rmic__tareas_moderadas'], df['calif__tareas_moderadas']= rmic, calificacion
            if calificacion=='Enfermedad comun':
                codigo_riesgo_tarea = 1
                posicion=find_indices(codigo_riesgo_tarea,'Td')
                if len(posicion)==0:
                    df['rmic__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
                    'No hay más tareas a evaluar'
                elif len(posicion)>0:
                    st.markdown('### Cálculo del RMic para tareas con riesgo leve')
                    res,calificacion,rmic=calculo_rmic(posicion,3.5,4)
                    df['rmic__tareas_leves'], df['calif__tareas_leves']= rmic, calificacion 
                    if calificacion=='Enfermedad comun':
                        'No hay más tareas a evaluar'
            else:
                df['rmic__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
                'No hay más tareas a evaluar'

    elif len(posicion)>0:
        st.markdown('### Cálculo del RMic para tareas con riesgo severo')
        res,calificacion,rmic=calculo_rmic(posicion,1.5,2) # Rango minimo: 1.5 ; Rango intermedio: 2
        df['rmic__tareas_severas'], df['calif__tareas_severas']= rmic, calificacion 
        if calificacion=='Enfermedad comun':
            codigo_riesgo_tarea = 2
            posicion=find_indices(codigo_riesgo_tarea,'Td')
            if len(posicion)==0:
                df['rmic__tareas_moderadas'], df['calif__tareas_moderadas']='No aplica', 'No aplica' 
                codigo_riesgo_tarea = 1
                posicion=find_indices(codigo_riesgo_tarea,'Td')
                if len(posicion)==0:
                    df['rmic__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
                    'No hay más tareas a evaluar'
                elif len(posicion)>0:
                    st.markdown('### Cálculo del RMic para tareas con riesgo leve')
                    res,calificacion,rmic=calculo_rmic(posicion,3.5,4)
                    df['rmic__tareas_leves'], df['calif__tareas_leves']= rmic, calificacion 
                    if calificacion=='Enfermedad comun':
                        'No hay más tareas a evaluar'
            elif len(posicion)>0:
                st.markdown('### Cálculo del RMic para tareas con riesgo moderado')
                res,calificacion,rmic=calculo_rmic(posicion,2.5,3)
                df['rmic__tareas_moderadas'], df['calif__tareas_moderadas']= rmic, calificacion
                if calificacion=='Enfermedad comun':
                    codigo_riesgo_tarea = 1
                    posicion=find_indices(codigo_riesgo_tarea,'Td')
                    if len(posicion)==0:
                        df['rmic__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
                        'No hay más tareas a evaluar'
                    elif len(posicion)>0:
                        st.markdown('### Cálculo del RMic para tareas con riesgo leve')
                        res,calificacion,rmic=calculo_rmic(posicion,3.5,4)
                        df['rmic__tareas_leves'], df['calif__tareas_leves']= rmic, calificacion  
                        if calificacion=='Enfermedad comun':
                            'No hay más tareas a evaluar'
                else:
                    df['rmic__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
                    'No hay más tareas a evaluar'
        else:
            df['rmic__tareas_moderadas'], df['calif__tareas_moderadas']='No aplica', 'No aplica' 
            df['rmic__tareas_leves'], df['calif__tareas_leves']='No aplica', 'No aplica'
            'No hay más tareas a evaluar'

# Botón de click para guardar calificación de la herramienta
guardar=st.button('Click para guardar registro')

if guardar==True: # Si se da click al botón, entonces se guardaran los datos en una planilla excel

    # Obtener fecha y hora en que se guarda el registro
    fecha_hora= datetime.now(pytz.timezone('Chile/Continental'))
    fecha_hora = fecha_hora.strftime("%d/%m/%Y %H:%M:%S")

    # Agregar campos de información del paciente, empresa, fecha del registro y el diagnóstico evaluado
    df.insert(1,'rut_paciente',rut_paciente)
    df.insert(2,'rut_empresa',rut_empresa) 
    df.insert(3,'fecha_informe',fecha_informe.strftime("%d/%m/%Y"))
    df.insert(4,'fecha_registro',fecha_hora)
    df.insert(5,'diagnostico',diagnostico) 

    # -------------Guardar en PC----------------
    # # Nombre del archivo donde se guardará el registro
    # file_path = 'registro_datos_EPME.xlsx'
    
    # # Definir dataframe que será guardado en el archivo
    # new_data = df.assign()

    # # Abrir archivo excel original
    # book = pd.read_excel(file_path, sheet_name='Sheet1')

    # # Pegar datos al archivo original
    # updated_data = book.append(new_data)

    # # Guardar resultados
    # with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    #     updated_data.to_excel(writer, sheet_name='Sheet1', index=False)

    # --------------Guardar en Sharepoint-------------------
    # Este código permite acceder a la carpeta en sharepoint y el archivo dónde se almacenan los registros de calificación
    server_url = "https://mutualcl.sharepoint.com/"
    site_url = server_url + "sites/SubgerenciaInnovacineInvestigacin"
    Username = 'estudiosyanalisis@mutual.cl'
    Password = 'Analytics2401#'
    Sharepoint_folder = 'Documentos Compartidos/Analytics/EP-ME/Registro Datos'
    nombre_archivo = 'registro_datos_EPME(V2).csv'

    authcookie = Office365(server_url, username = Username, password=Password).GetCookies()
    site = Site(site_url, version=Version.v365, authcookie=authcookie)
    folder = site.Folder(Sharepoint_folder)
    data=StringIO(str(folder.get_file(nombre_archivo),'utf-8'))
    df_original=pd.read_csv(data,sep=';')
 
    # Se pegan los últimos registros al archivo original
    new_data = df.assign()
    df_final = df_original.append(new_data)
    df_final.to_csv(nombre_archivo,sep=';',index=False)

    # Se carga la información en sharepoint
    with open(nombre_archivo, mode='rb') as file:
        fileContent = file.read()
    folder.upload_file(fileContent, nombre_archivo)
    file.close()

    mensaje='Registro guardado!'
    txt='<p style="font-family:Courier; color: Red; font-size: 15px">'+mensaje+'</p>'
    st.markdown(txt, unsafe_allow_html=True)

# Desplegar logo Mutual al final de la pagina
c1,c2,c3=st.columns(3)
with c1:
    ''
    ''
    ''
    st.image(image, width=400,caption='Subgerencia de Innovación e Investigación')


