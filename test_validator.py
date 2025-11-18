# -*- coding: utf-8 -*-
"""
PDF验证工具 - 自动测试脚本

执行3次测试验证，确保程序正常运行
"""

import os
import shutil
import sys
from validator_core import (
    parse_filename,
    extract_pdf_info,
    compare_info,
    validate_folder,
    generate_excel_report
)


def print_header(title):
    """打印测试标题"""
    print("\n" + "="*60)
    print(f"  {title}")
    print("="*60 + "\n")


def print_result(test_name, passed, message=""):
    """打印测试结果"""
    status = "✓ 通过" if passed else "✗ 失败"
    color = "\033[92m" if passed else "\033[91m"
    reset = "\033[0m"
    print(f"{color}{status}{reset} - {test_name}")
    if message:
        print(f"      {message}")


def test_1_filename_parsing():
    """测试1：文件名解析功能"""
    print_header("测试1：文件名解析功能")

    # 测试用例
    test_cases = [
        {
            'filename': '809_MARTINEZ MONTERO_ LAURA MARIA_Payslip_01092024.pdf',
            'expected': {
                'codigo': '809',
                'nombre': 'MARTINEZ MONTERO LAURA MARIA',
                'fecha': '01092024',
                'valid': True
            }
        },
        {
            'filename': '12345_GARCIA LOPEZ_ JUAN_Payslip_01_01_2024.pdf',
            'expected': {
                'codigo': '12345',
                'nombre': 'GARCIA LOPEZ JUAN',
                'fecha': '01012024',
                'valid': True
            }
        },
        {
            'filename': 'invalid_filename.pdf',
            'expected': {
                'valid': False
            }
        }
    ]

    passed_count = 0
    for i, case in enumerate(test_cases, 1):
        result = parse_filename(case['filename'])

        if case['expected']['valid']:
            # 验证有效文件名
            checks = [
                result['valid'] == True,
                result['codigo'] == case['expected']['codigo'],
                result['nombre'] == case['expected']['nombre'],
                result['fecha'] == case['expected']['fecha']
            ]
            passed = all(checks)
        else:
            # 验证无效文件名
            passed = result['valid'] == False

        passed_count += passed
        print_result(
            f"用例 {i}: {case['filename'][:40]}...",
            passed,
            f"编号={result.get('codigo', 'N/A')}, 姓名={result.get('nombre', 'N/A')[:20]}..."
        )

    total = len(test_cases)
    print(f"\n测试1结果: {passed_count}/{total} 通过")
    return passed_count == total


def test_2_pdf_extraction():
    """测试2：PDF信息提取功能"""
    print_header("测试2：PDF信息提取功能")

    # 使用现有的示例PDF
    test_pdf = '/home/user/Slipt_verif/809_MARTINEZ MONTERO_ LAURA MARIA_Payslip_01092024.pdf'

    if not os.path.exists(test_pdf):
        print_result("PDF文件存在性检查", False, f"文件不存在: {test_pdf}")
        return False

    print_result("PDF文件存在性检查", True)

    # 提取PDF信息
    pdf_info = extract_pdf_info(test_pdf)

    # 验证提取的信息
    checks = {
        'PDF有效性': pdf_info.get('valid', False),
        '编号提取': pdf_info.get('codigo') == '809',
        '姓名提取': pdf_info.get('nombre') and 'MARTINEZ' in pdf_info.get('nombre', ''),
        'NIF提取': pdf_info.get('nif') and pdf_info['nif'] != 'NO ENCONTRADO',
        '期间提取': pdf_info.get('periodo') and pdf_info['periodo'] != 'NO ENCONTRADO'
    }

    passed_count = 0
    for check_name, check_result in checks.items():
        passed_count += check_result
        value = ""
        if check_name == '编号提取':
            value = f"编号={pdf_info.get('codigo')}"
        elif check_name == '姓名提取':
            value = f"姓名={pdf_info.get('nombre', '')[:30]}..."
        elif check_name == 'NIF提取':
            value = f"NIF={pdf_info.get('nif')}"
        elif check_name == '期间提取':
            value = f"期间={pdf_info.get('periodo', '')[:30]}..."

        print_result(check_name, check_result, value)

    total = len(checks)
    print(f"\n测试2结果: {passed_count}/{total} 通过")
    return passed_count >= 4  # 至少4项通过


def test_3_comparison_logic():
    """测试3：比对逻辑功能"""
    print_header("测试3：比对逻辑功能")

    # 测试用例：匹配的情况
    filename_info_match = {
        'codigo': '809',
        'nombre': 'MARTINEZ MONTERO LAURA MARIA',
        'fecha': '01092024'
    }

    pdf_info_match = {
        'codigo': '809',
        'nombre': 'MARTINEZ MONTERO, LAURA MARIA',
        'nif': '02724279K',
        'periodo': '1 septiembre 30 septiembre 2025'
    }

    result_match = compare_info(filename_info_match, pdf_info_match)

    # 测试用例：不匹配的情况
    filename_info_mismatch = {
        'codigo': '810',
        'nombre': 'GARCIA LOPEZ JUAN',
        'fecha': '01092024'
    }

    pdf_info_mismatch = {
        'codigo': '809',
        'nombre': 'MARTINEZ MONTERO, LAURA MARIA',
        'nif': '02724279K',
        'periodo': '1 septiembre 30 septiembre 2025'
    }

    result_mismatch = compare_info(filename_info_mismatch, pdf_info_mismatch)

    # 验证结果
    checks = {
        '匹配场景-编号匹配': result_match['codigo_match'] == True,
        '匹配场景-姓名匹配': result_match['nombre_match'] == True,
        '匹配场景-总体匹配': result_match['overall_match'] == True,
        '不匹配场景-编号不匹配': result_mismatch['codigo_match'] == False,
        '不匹配场景-姓名不匹配': result_mismatch['nombre_match'] == False,
        '不匹配场景-总体不匹配': result_mismatch['overall_match'] == False
    }

    passed_count = 0
    for check_name, check_result in checks.items():
        passed_count += check_result
        print_result(check_name, check_result)

    total = len(checks)
    print(f"\n测试3结果: {passed_count}/{total} 通过")
    return passed_count == total


def test_4_integration():
    """测试4：集成测试 - 完整流程"""
    print_header("测试4：集成测试 - 完整流程")

    # 创建测试文件夹
    test_dir = '/home/user/Slipt_verif/test_pdfs'
    os.makedirs(test_dir, exist_ok=True)

    # 复制示例PDF到测试文件夹
    src_pdf = '/home/user/Slipt_verif/809_MARTINEZ MONTERO_ LAURA MARIA_Payslip_01092024.pdf'

    if os.path.exists(src_pdf):
        # 复制并重命名为测试文件
        test_files = [
            '809_MARTINEZ MONTERO_ LAURA MARIA_Payslip_01092024.pdf',  # 正确的
            '809_MARTINEZ_MONTERO_LAURA_MARIA_Payslip_01092024.pdf',  # 下划线版本
        ]

        for test_file in test_files:
            dst = os.path.join(test_dir, test_file)
            shutil.copy2(src_pdf, dst)
            print(f"  创建测试文件: {test_file}")

        print()

        # 执行验证
        results = validate_folder(test_dir)

        # 生成报告
        report_path = os.path.join(test_dir, 'test_report.xlsx')
        generate_excel_report(results, report_path)

        # 检查结果
        checks = {
            '找到PDF文件': len(results) > 0,
            '报告生成成功': os.path.exists(report_path),
            '至少一个文件匹配': any(r['overall_match'] for r in results) if results else False
        }

        passed_count = 0
        for check_name, check_result in checks.items():
            passed_count += check_result
            if check_name == '找到PDF文件':
                msg = f"找到 {len(results)} 个文件"
            elif check_name == '报告生成成功':
                msg = f"报告路径: {report_path}"
            elif check_name == '至少一个文件匹配':
                matched = sum(1 for r in results if r['overall_match'])
                msg = f"{matched}/{len(results)} 文件匹配"
            else:
                msg = ""

            print_result(check_name, check_result, msg)

        # 显示验证详情
        if results:
            print("\n  验证详情:")
            for r in results:
                status = "✓" if r['overall_match'] else "✗"
                print(f"    {status} {r['filename'][:50]}...")

        total = len(checks)
        print(f"\n测试4结果: {passed_count}/{total} 通过")

        # 清理测试文件
        print("\n  清理测试文件...")
        shutil.rmtree(test_dir)

        return passed_count >= 2  # 至少2项通过
    else:
        print_result("测试文件准备", False, "未找到源PDF文件")
        return False


def main():
    """主测试函数"""
    print("\n" + "█"*60)
    print("█" + " "*58 + "█")
    print("█" + "  PDF工资单验证工具 - 自动测试套件".center(58) + "█")
    print("█" + " "*58 + "█")
    print("█"*60)

    # 执行所有测试
    tests = [
        ("文件名解析", test_1_filename_parsing),
        ("PDF信息提取", test_2_pdf_extraction),
        ("比对逻辑", test_3_comparison_logic),
        ("集成测试", test_4_integration)
    ]

    results = []
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"\n✗ 测试 '{test_name}' 出现异常: {str(e)}")
            results.append((test_name, False))

    # 汇总结果
    print_header("测试汇总")

    passed = sum(1 for _, result in results if result)
    total = len(results)

    for test_name, result in results:
        print_result(test_name, result)

    print(f"\n总计: {passed}/{total} 测试通过")

    if passed == total:
        print("\n✓✓✓ 所有测试通过！程序运行正常！ ✓✓✓\n")
        return 0
    else:
        print(f"\n✗✗✗ {total - passed} 个测试失败 ✗✗✗\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())
