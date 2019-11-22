pipeline {
  agent any
  stages {
    stage('Paso 1') {
      parallel {
        stage('Paso 1') {
          steps {
            pwd(tmp: true)
            echo 'Hola Mundo'
            readFile 'permisos.ps1'
          }
        }

        stage('Paso 2') {
          agent any
          steps {
            readFile 'permisos.ps1'
            fileExists 'permisos.ps1'
            writeFile(file: 'permisos.log', text: 'El fichero existe en la rama')
          }
        }

      }
    }

  }
}