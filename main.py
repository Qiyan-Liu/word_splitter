#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word文档拆分工具 - 主程序入口

作者: QiyanLiu
日期: 2025-08-06

这个脚本提供了Word文档拆分工具的主要入口点。
用户可以直接在这里修改配置参数，无需额外的配置文件。

配置说明：
- INPUT_DIR: 输入目录，包含要拆分的Word文档
- OUTPUT_DIR: 输出目录，拆分后的文档将保存在这里
- MIN_LEVEL: 最小拆分层级（1-6，对应Word的标题1-6）
- FILE_THREAD_COUNT: 处理多个文档的线程数
- CHAPTER_THREAD_COUNT: 处理单个文档章节的线程数

性能优化建议：
1. 少量大文档：建议 FILE_THREAD_COUNT=1-2, CHAPTER_THREAD_COUNT=2-4
2. 大量小文档：建议 FILE_THREAD_COUNT=4-8, CHAPTER_THREAD_COUNT=1-2
3. SSD硬盘：可使用更高的FILE_THREAD_COUNT
4. 机械硬盘：建议降低FILE_THREAD_COUNT避免磁盘竞争
5. 内存充足：可增加CHAPTER_THREAD_COUNT
6. 内存有限：建议保守的CHAPTER_THREAD_COUNT配置
"""

import os
import sys
import time
from pathlib import Path
from word_splitter import DocumentProcessor

try:
    from tqdm import tqdm
except ImportError:
    # 如果没有tqdm，使用简单的进度显示
    class tqdm:
        def __init__(self, iterable=None, total=None, desc="", unit=""):
            self.iterable = iterable
            self.total = total or (len(iterable) if iterable else 0)
            self.desc = desc
            self.current = 0
            if desc:
                print(f"{desc}: 0/{self.total}")
        
        def update(self, n=1):
            self.current += n
            if self.desc:
                print(f"\r{self.desc}: {self.current}/{self.total}", end="", flush=True)
        
        def close(self):
            if self.desc:
                print()  # 换行
        
        def __enter__(self):
            return self
        
        def __exit__(self, *args):
             self.close()

def process_documents_with_progress(processor, doc_files, pbar):
    """带进度条的文档处理函数"""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    
    # 使用线程池处理多个文档
    with ThreadPoolExecutor(max_workers=processor.file_thread_count) as executor:
        futures = [executor.submit(process_single_document_with_callback, 
                                 processor, doc_file, pbar) 
                  for doc_file in doc_files]
        
        for future in as_completed(futures):
            try:
                result = future.result()
                # 进度条已在回调中更新
            except Exception as e:
                print(f"\n❌ 文档处理失败: {e}")
                pbar.update(1)

def process_single_document_with_callback(processor, doc_path, pbar):
    """处理单个文档并更新进度条"""
    try:
        # 分析文档结构
        doc, chapters = processor.splitter.analyze_document_structure(str(doc_path))
        
        if not chapters:
            pbar.set_description(f"📄 跳过: {doc_path.name} (无章节)")
            pbar.update(1)
            return f"跳过: {doc_path.name}"
        
        # 更新进度条描述
        pbar.set_description(f"📄 处理: {doc_path.name} ({len(chapters)}章节)")
        
        # 创建文档专用的输出目录
        doc_output_dir = processor.output_dir / doc_path.stem
        doc_output_dir.mkdir(parents=True, exist_ok=True)
        
        # 使用线程池处理章节
        from concurrent.futures import ThreadPoolExecutor, as_completed
        with ThreadPoolExecutor(max_workers=processor.chapter_thread_count) as executor:
            futures = [executor.submit(processor.splitter.create_chapter_document, 
                                     doc, chapter, str(doc_output_dir))
                      for chapter in chapters]
            
            created_files = []
            for future in as_completed(futures):
                try:
                    result = future.result()
                    created_files.append(result)
                except Exception as e:
                    print(f"\n❌ 章节处理失败: {e}")
        
        # 更新进度条
        pbar.set_description(f"✅ 完成: {doc_path.name} ({len(created_files)}个文件)")
        pbar.update(1)
        
        return f"完成: {doc_path.name} -> {len(created_files)} 个章节"
        
    except Exception as e:
        pbar.set_description(f"❌ 失败: {doc_path.name}")
        pbar.update(1)
        return f"失败: {doc_path.name}"

def main():
    """主函数 - 配置参数并启动文档处理"""
    
    # ==================== 配置参数 ====================
    # 用户可以直接在这里修改以下参数
    
    # 输入输出目录
    INPUT_DIR = "input"          # 输入目录
    OUTPUT_DIR = "output"        # 输出目录
    
    # 文档拆分配置
    MIN_LEVEL = 5               # 最小拆分层级（1-6）
    
    # 多线程配置
    FILE_THREAD_COUNT = 4       # 处理多个文档的线程数（建议1-8）
    CHAPTER_THREAD_COUNT = 4    # 处理单个文档章节的线程数（建议1-4）
    
    # ==================== 参数验证 ====================
    
    # 验证目录路径
    if not os.path.exists(INPUT_DIR):
        print(f"错误：输入目录 '{INPUT_DIR}' 不存在")
        return
    
    # 验证参数范围
    if not (1 <= MIN_LEVEL <= 6):
        print(f"警告：MIN_LEVEL应在1-6之间，当前值：{MIN_LEVEL}，使用默认值3")
        MIN_LEVEL = 3
    
    if not (1 <= FILE_THREAD_COUNT <= 16):
        print(f"警告：FILE_THREAD_COUNT应在1-16之间，当前值：{FILE_THREAD_COUNT}，使用默认值2")
        FILE_THREAD_COUNT = 2
    
    if not (1 <= CHAPTER_THREAD_COUNT <= 8):
        print(f"警告：CHAPTER_THREAD_COUNT应在1-8之间，当前值：{CHAPTER_THREAD_COUNT}，使用默认值2")
        CHAPTER_THREAD_COUNT = 2
    
    # 创建输出目录
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # ==================== 启动处理 ====================
    
    print("\n" + "=" * 60)
    print("📄 Word文档拆分工具")
    print("=" * 60)
    print(f"📁 输入目录: {INPUT_DIR}")
    print(f"📂 输出目录: {OUTPUT_DIR}")
    print(f"📊 最小拆分层级: {MIN_LEVEL}")
    print(f"🧵 文档处理线程数: {FILE_THREAD_COUNT}")
    print(f"⚡ 章节处理线程数: {CHAPTER_THREAD_COUNT}")
    print(f"🚀 总线程数: {FILE_THREAD_COUNT * CHAPTER_THREAD_COUNT}")
    print("-" * 60)
    
    try:
        # 检查输入文档
        input_path = Path(INPUT_DIR)
        doc_files = list(input_path.glob('*.docx')) + list(input_path.glob('*.doc'))
        
        if not doc_files:
            print(f"❌ 在 {INPUT_DIR} 中未找到Word文档")
            return
        
        print(f"\n🔍 发现 {len(doc_files)} 个文档:")
        for i, doc_file in enumerate(doc_files, 1):
            print(f"   {i}. {doc_file.name}")
        
        print(f"\n🚀 开始处理...")
        
        # 创建文档处理器
        processor = DocumentProcessor(
            input_dir=INPUT_DIR,
            output_dir=OUTPUT_DIR,
            file_thread_count=FILE_THREAD_COUNT,
            chapter_thread_count=CHAPTER_THREAD_COUNT,
            min_level=MIN_LEVEL
        )
        
        # 使用进度条处理文档
        with tqdm(total=len(doc_files), desc="📄 处理文档", unit="个") as pbar:
            start_time = time.time()
            
            # 重写处理方法以支持进度条
            process_documents_with_progress(processor, doc_files, pbar)
            
            end_time = time.time()
            elapsed_time = end_time - start_time
        
        print(f"\n✅ 处理完成！")
        print(f"⏱️  总耗时: {elapsed_time:.2f} 秒")
        print(f"📂 结果保存在: {os.path.abspath(OUTPUT_DIR)}")
        print("=" * 60)
        
    except KeyboardInterrupt:
        print("\n⏹️  用户中断操作")
    except Exception as e:
        print(f"\n❌ 处理过程中发生错误: {e}")
        print("📋 详细错误信息已记录到 word_splitter.log")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()