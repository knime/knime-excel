#!groovy
def BN = (BRANCH_NAME == 'master' || BRANCH_NAME.startsWith('releases/')) ? BRANCH_NAME : 'releases/2025-12'

library "knime-pipeline@$BN"

properties([
    pipelineTriggers([
        upstream('knime-base/' + env.BRANCH_NAME.replaceAll('/', '%2F'))
	]),
    parameters(workflowTests.getConfigurationsAsParameters() + fsTests.getFSConfigurationsAsParameters()),
    buildDiscarder(logRotator(numToKeepStr: '5')),
    disableConcurrentBuilds()
])

SSHD_IMAGE = "${dockerTools.ECR}/knime/sshd:alpine3.11"

try {
    knimetools.defaultTychoBuild('org.knime.update.ext.poi')

    configs = [
        "Workflowtests" : {
            workflowTests.runTests (
                dependencies: [
                    repositories: [
                        "knime-aws",
                        "knime-excel",
                        "knime-timeseries",
                        "knime-jep",
                        "knime-datageneration",
                        "knime-filehandling",
                        "knime-jfreechart",
                        "knime-distance",
                        "knime-exttool",
                        "knime-chemistry",
                        "knime-python",
                        "knime-js-core",
                        "knime-js-base",
                        "knime-cloud",
                        "knime-textprocessing",
                        "knime-database",
                        "knime-kerberos",
                        "knime-streaming",
                        "knime-xml"
                        ]
                ],
                sidecarContainers: [
                    [ image: SSHD_IMAGE, namePrefix: "SSHD", port: 22 ]
                ],
            )
        },
        "Filehandlingtests" : {
            workflowTests.runFilehandlingTests (
                dependencies: [
                    repositories: [
                        "knime-datageneration",
                        "knime-excel",
                        "knime-expressions",
                        "knime-jep",
                        "knime-js-base",
                        "knime-streaming",
                        "knime-xml"
                    ]
                ],
            )
        }
    ]

    parallel configs

    stage('Sonarqube analysis') {
        env.lastStage = env.STAGE_NAME
        workflowTests.runSonar()
	}
} catch (ex) {
    currentBuild.result = 'FAILURE'
    throw ex
} finally {
    notifications.notifyBuild(currentBuild.result);
}
/* vim: set shiftwidth=4 expandtab smarttab: */
