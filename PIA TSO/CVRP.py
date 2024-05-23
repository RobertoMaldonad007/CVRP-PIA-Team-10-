import re
import math
import random

from time import process_time

import sys
import os
import xlsxwriter

def DistanciaEuclidiana(Nodo_i, Nodo_j):
    return math.sqrt(pow((Nodo_i[1] - Nodo_j[1]), 2) + pow((Nodo_i[2] - Nodo_j[2]), 2))

def BusquedaLocal(Vehiculos):
    ListaTabu = []
    Total = 0
    
    NuevasRutas = []
    
    for Vehiculo in Vehiculos:
        DistanciaCorta = Vehiculo.DistanciaRecorrida
        RutaActual = Vehiculo.Ruta[1:-1]
        Nodo_i = 0
        Nodo_j = 0
        NoMejora = 0
        for _ in range (len(RutaActual)):
            i = 0
            while (i < (len(Vehiculo.Ruta) - 2)):
                DistanciaTotal = 0
                NodoAnterior = 0
                
                if i == (len(RutaActual) - 1):
                    RutaActual[i], RutaActual[0] = RutaActual[0], RutaActual[i]
                else:
                    RutaActual[i], RutaActual[(i + 1)] = RutaActual[(i + 1)], RutaActual[i]
                
                for x in RutaActual:
                    Distancia = round(DistanciaEuclidiana(Puntos[NodoAnterior], Puntos[(x - 1)]))
                    DistanciaTotal = (DistanciaTotal + Distancia)
                    NodoAnterior = (x - 1)
                Distancia = round(DistanciaEuclidiana(Puntos[(NodoAnterior - 1)], Puntos[0]))
                DistanciaTotal = (DistanciaTotal + Distancia)
                
                if DistanciaTotal < DistanciaCorta:
                    NoMejora = 0
                    if i == (len(RutaActual) - 1):
                        Nodo_i = RutaActual[0]
                        Nodo_j = RutaActual[i]
                        
                        DistanciaCorta = DistanciaTotal
                        MejorRuta = RutaActual
                    else:
                        Nodo_i = RutaActual[(i + 1)]
                        Nodo_j = RutaActual[i]
                        
                        DistanciaCorta = DistanciaTotal
                        MejorRuta = RutaActual
                        
                if i == (len(RutaActual) - 1):
                    RutaActual[0], RutaActual[i] = RutaActual[i], RutaActual[0]
                else:
                    RutaActual[(i + 1)], RutaActual[i] = RutaActual[i], RutaActual[(i + 1)]
                i = (i + 1)
            
            if (MejorRuta == []):
                NoMejora = (NoMejora + 1)
                MejorRuta = RutaActual
            else:
                if Nodo_i != 0 and Nodo_j != 0:
                    if len(ListaTabu) == 0:
                        NoMejora = 0
                        I = MejorRuta.index(Nodo_i)
                        J = MejorRuta.index(Nodo_j)
                        MejorRuta[I], MejorRuta[J] = MejorRuta[J], MejorRuta[I]
                        ListaTabu.append([Nodo_i, Nodo_j])
                    else:
                        if [Nodo_i, Nodo_j] != ListaTabu[-1]:
                            NoMejora = 0
                            I = MejorRuta.index(Nodo_i)
                            J = MejorRuta.index(Nodo_j)
                            MejorRuta[I], MejorRuta[J] = MejorRuta[J], MejorRuta[I]
                            ListaTabu.append([Nodo_i, Nodo_j])
                        else:
                            NoMejora = (NoMejora + 1)
            if NoMejora > 1:
                break
        MejorRuta.insert(0, 1)
        MejorRuta.append(1)
        
        Ruta = [Vehiculo.Numero, MejorRuta, DistanciaCorta, Vehiculo.CapacidadTotal]
        NuevasRutas.append(Ruta)
        
        MejorRuta = []
        ListaTabu.clear()
    return NuevasRutas

def RandomizedConstructiveHueristic(Datos, Dimension, NumeroDeVehiculos, Q):
    Rutas = []
    NodosVisitados = []
    
    for i in range(NumeroDeVehiculos):
        TodasLasRutas = []
        Distancias = []
        Capacidades = []
        Nodos = []
        
        for x in range(NumeroDeVehiculos):
            Soluciones = [Datos[0]]
            Demanda = 0
            DistanciaTotal = 0
            
            while True:
                Aleatorio = random.choice(Datos)
                
                if Aleatorio[0] != 1:
                    if Aleatorio[0] not in NodosVisitados:
                        if (Aleatorio not in Soluciones) and (Demanda + Aleatorio[3] < Q):
                            Distancia = round(DistanciaEuclidiana(Soluciones[-1], Aleatorio))
                            Soluciones.append(Aleatorio)
                            DistanciaTotal = round(DistanciaTotal + Distancia)
                            Demanda = Demanda + Aleatorio[3]
                else:
                    if len(Soluciones) >= 2:
                        break
            
            Distancia = round(DistanciaEuclidiana(Soluciones[-1], Datos[0]))
            DistanciaTotal = round(DistanciaTotal + Distancia)
            Soluciones.append(Datos[0])
            TodasLasRutas.append(Soluciones)
            Distancias.append(DistanciaTotal)
            Capacidades.append(Demanda)
            Soluciones = []
        
        MejorRuta = Distancias.index(min(Distancias))
        
        for Nodo in TodasLasRutas[MejorRuta]:
            Nodos.append(Nodo[0:1][0])
            if (Nodo[0:1][0] != 1):
                NodosVisitados.append(Nodo[0:1][0])
        
        Ruta = [(i + 1), Nodos, min(Distancias), Capacidades[MejorRuta]]
        Rutas.append(Ruta)
    return Rutas

class Vehiculo:
    def __init__(self, Numero, Ruta, DistanciaRecorrida, CapacidadTotal):
        self.Numero = Numero
        self.Ruta = Ruta
        self.DistanciaRecorrida = DistanciaRecorrida
        self.CapacidadTotal = CapacidadTotal

if __name__=="__main__":
    Instancia = input("Instance name: ")
    
    InicioDeEjecucion = process_time()
    
    try:
        with open(Instancia, "r") as Archivo:
            Datos = Archivo.read().splitlines()
            
            Dimension, NumeroDeRutas = ([int(i) for i in re.findall(r'\d+', Instancia)])
            
            Capacidad = ([int(i) for i in re.findall(r'\d+', Datos[5])])
            Capacidad = Capacidad[0]

            del Datos[0:7]
            
            Puntos = []
            for Linea in Datos:
                if Linea.strip() != "DEMAND_SECTION":
                    Columnas = Linea.split()
                    Puntos.append([int(Columnas[0])] + [int(Columnas[1])] + [int(Columnas[2])])
                else:
                    del Datos[0:Dimension]
                    break

            for Linea in Datos:
                if Linea.strip() != "DEMAND_SECTION":
                    if Linea.strip() == "DEPOT_SECTION":
                        break
                    else:
                        Columnas = Linea.split()
                        Puntos[(int(Columnas[0]) - 1)].append(int(Columnas[1]))

            NodosVisitadosPorRuta = []
            NodosVisitados = []
            Rutas = []
            
            # Constructive heuristic:
            for i in range(NumeroDeRutas):
                NumeroRutaActual = (i + 1)
                CapacidadTotal = 0
                DistanciaTotal = 0
                
                for _ in range(Dimension):
                    DistanciaMasLejana = 0
                    
                    for j in range(Dimension):
                        if Puntos[i][0] != Puntos[j][0] and Puntos[j][0] != 1:
                            if Puntos[j][0] not in NodosVisitados:
                                Distancia_i_j = round(DistanciaEuclidiana(Puntos[i], Puntos[j]))
                                if (Distancia_i_j > DistanciaMasLejana) and (CapacidadTotal + Puntos[j][3]) < Capacidad:
                                    DistanciaMasLejana = Distancia_i_j
                                    
                                    Nodo = Puntos[j][0]
                                    CapacidadDelNodo = Puntos[j][3]
                    
                    if DistanciaMasLejana != 0:
                        CapacidadTotal = CapacidadTotal + CapacidadDelNodo
                        NodosVisitadosPorRuta.append(Nodo)
                        DistanciaTotal = round(float(str(DistanciaTotal + DistanciaMasLejana)[:6]))
                        i = (Nodo - 1)
                        NodosVisitados.append(Nodo)
                NodosVisitadosPorRuta.insert(0, 1)
                NodosVisitadosPorRuta.append(1)
                
                DistanciaTotal = round(float(str(DistanciaTotal + DistanciaEuclidiana(Puntos[(Nodo - 1)], Puntos[0]))[:6]))
                
                VehiculoActual = Vehiculo(NumeroRutaActual, NodosVisitadosPorRuta[0:], DistanciaTotal, CapacidadTotal)
                Rutas.append(VehiculoActual)
                
                NodosVisitadosPorRuta.clear()

        try:
            SolucionInstancia = Instancia.replace("vrp", "sol")
            with open(SolucionInstancia, "r") as ArchivoSolucion:
                Solucion = ArchivoSolucion.read().splitlines()
            
                SolucionOptima = ([int(i) for i in re.findall(r'\d+', Solucion[-1])])
        except FileNotFoundError:
            print("Optimal solution for instance not found.")

        RutasLS = BusquedaLocal(Rutas)
        
        RutasRandom = RandomizedConstructiveHueristic(Puntos, Dimension, NumeroDeRutas, Capacidad)
        
        Costos = []
        
        Tabla = xlsxwriter.Workbook('Tabla.xlsx')
        Hoja = Tabla.add_worksheet()
        
        Formato = Tabla.add_format({'border': 1})
        Formato.set_align('center')
        
        EstiloRuta = Tabla.add_format({'border': 1})
        
        Hoja.write('A1', 'Instance:')
        Hoja.write('B1', Instancia)
        Hoja.merge_range('A3:C3', 'Constructive heuristic:', Formato)
        Hoja.write(3, 0, 'Truck:', Formato)
        Hoja.write(3, 1, 'Route:', Formato)
        Hoja.write(3, 2, 'Capacity:', Formato)
        
        i = 4
        CostoCH = 0
        for CadaRuta in Rutas:
            Hoja.write(i, 0, CadaRuta.Numero, Formato)
            Hoja.write(i, 1, ''.join(str(CadaRuta.Ruta)), EstiloRuta)
            Hoja.write(i, 2, CadaRuta.CapacidadTotal, Formato)
            CostoCH = CostoCH + CadaRuta.DistanciaRecorrida
            
            i = (i + 1)
        
        Costos.append(CostoCH)
        
        Hoja.write(i, 3, 'Cost:', Formato)
        Hoja.write(i, 4, CostoCH, Formato)
        
        i = (i + 2)
        Hoja.write(i, 0,'Local search:', Formato)
        
        i = (i + 1)
        Hoja.write(i, 0, 'Truck:', Formato)
        Hoja.write(i, 1, 'Route:', Formato)
        Hoja.write(i, 2, 'Capacity:', Formato)
        
        i = (i + 1)
        CostoLS = 0
        for CadaRuta in RutasLS:
            Hoja.write(i, 0, CadaRuta[0], Formato)
            Hoja.write(i, 1, ''.join(str(CadaRuta[1])), EstiloRuta)
            Hoja.write(i, 2, CadaRuta[3], Formato)
            CostoLS = CostoLS + CadaRuta[2]
            
            i = (i + 1)
        
        Hoja.write(i, 3, 'Cost:', Formato)
        Hoja.write(i, 4, CostoLS, Formato)
        
        Costos.append(CostoLS)
        
        i = (i + 2)
        Hoja.write(i, 0,'Randomized constructive hueristic:', Formato)
        
        i = (i + 1)
        Hoja.write(i, 0, 'Truck:', Formato)
        Hoja.write(i, 1, 'Route:', Formato)
        Hoja.write(i, 2, 'Capacity:', Formato)
        
        i = (i + 1)
        CostoR = 0
        for CadaRuta in RutasRandom:
            Hoja.write(i, 0, CadaRuta[0], Formato)
            Hoja.write(i, 1, ''.join(str(CadaRuta[1])), EstiloRuta)
            Hoja.write(i, 2, CadaRuta[3], Formato)
            CostoR = CostoR + CadaRuta[2]
            
            i = (i + 1)
        
        Hoja.write(i, 3, 'Cost:', Formato)
        Hoja.write(i, 4, CostoR, Formato)
        
        Costos.append(CostoR)
        
        Hoja.merge_range('G12:H12', 'Best known solution:', Formato)
        Hoja.merge_range('G13:H13', min(Costos), Formato)
        
        Hoja.merge_range('I12:J12', 'Optimal solution (given by autors):', Formato)
        Hoja.merge_range('I13:J13', SolucionOptima[0], Formato)
        
        Hoja.write('K12', 'Deviation:', Formato)
        Hoja.write('K13', (str(100*((min(Costos) - SolucionOptima[0])/SolucionOptima[0]))[:5]) + '%', Formato)
        
        Hoja.autofit()
        Tabla.close()
        
        os.startfile('Tabla.xlsx')
        
        FinDeEjecucion = process_time()
        
        print(f"\nTotal program execution time: {(FinDeEjecucion - InicioDeEjecucion)} seconds")
        
    except FileNotFoundError:
        print("Instance not found.")