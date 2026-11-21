# Despliegue Automatizado con Kubernetes y Argo CD (GitOps)

Este repositorio contiene la configuración declarativa para el despliegue automatizado de una aplicación web en Kubernetes siguiendo las prácticas de **GitOps** mediante **Argo CD**.

## Arquitectura de la Solución

El repositorio actúa como la **única fuente de verdad** (Single Source of Truth). Argo CD monitorea continuamente este repositorio y reconcilia cualquier desviación con el clúster de Kubernetes mediante sincronización automática (*Automated Sync*).

### Manifiestos incluidos:
- **`deployment.yml`**: Define el despliegue de la aplicación web (`registry.k8s.io/hpa-example`) con 2 réplicas iniciales, configuración de etiquetas y límites estrictos de CPU y memoria (`requests` y `limits`) requeridos para el escalamiento.
- **`service.yml`**: Expone la aplicación hacia el exterior mediante un servicio de tipo `NodePort` en el puerto `30080` (mapeado al puerto `80` del contenedor).
- **`hpa.yml`**: Implementa el auto-escalado horizontal (`HorizontalPodAutoscaler`) configurado para mantener un mínimo de 2 y un máximo de 5 réplicas, con un umbral de utilización promedio de CPU del 50%.
