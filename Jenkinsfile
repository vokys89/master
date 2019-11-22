pipeline {
  agent any
  stages {
    stage('Paso 1') {
      parallel {
        stage('Paso 1') {
          steps {
            powershell(script: 'permisos.ps1', label: 'Permisos', returnStatus: true, returnStdout: true)
          }
        }

        stage('Paso 2') {
          steps {
            readFile 'permisos.ps1'
          }
        }

      }
    }

  }
}