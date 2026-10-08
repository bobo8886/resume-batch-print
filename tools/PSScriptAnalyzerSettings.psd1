@{
    # PSScriptAnalyzer 配置
    # 目标：只报"真的会有问题"的规则，不要把风格偏好变成 CI 红叉。

    Severity = @('Error', 'Warning')

    ExcludeRules = @(
        # 这是给人双击运行的交互式工具，控制台彩色输出是刻意为之
        'PSAvoidUsingWriteHost'

        # 内部辅助函数用的是自定义动词 / 复数名词，不是为了发布成模块
        'PSUseApprovedVerbs'
        'PSUseSingularNouns'

        # 脚本本身就是"一键运行"的形态，不需要 SupportsShouldProcess
        'PSUseShouldProcessForStateChangingFunctions'

        # 有大量合法的空 catch（清理 / 兜底），故意忽略
        'PSAvoidUsingEmptyCatchBlock'

        # 不是模块，不需要清单字段
        'PSUseToExportFieldsInManifest'
        'PSUseBOMForUnicodeEncodedFile'
    )
}
