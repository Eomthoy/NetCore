using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.StaticFiles;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace Core.Common.Helper
{
    /// <summary>
    /// 文件上传、下载
    /// </summary>
    public class FileHelper
    {
        /// <summary>
        /// 文件上传
        /// </summary>
        /// <param name="path">存放地址</param>
        /// <param name="files">所有文件</param>
        /// <param name="maxSize">最大上传大小（默认10M）</param>
        public static void Upload(string path, IFormFileCollection files, int maxSize = 10)
        {
            try
            {
                if (files != null && files.Sum(x => x.Length) > 0)
                {
                    foreach (IFormFile file in files)
                    {
                        Upload(path, file, maxSize);
                    }
                }
                else
                {
                    throw new Exception("上传文件不能为空");
                }
            }
            catch (Exception ex)
            {
                while (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                }
                throw ex;
            }
        }
        /// <summary>
        /// 文件上传
        /// </summary>
        /// <param name="path">存放地址</param>
        /// <param name="file">单个文件</param>
        /// <param name="maxSize">最大上传大小（默认10M）</param>
        public static void Upload(string path, IFormFile file, int maxSize = 10)
        {
            try
            {
                if (file != null && file.Length > 0)
                {
                    if (file.Length > 1024 * 1024 * maxSize)
                    {
                        throw new Exception($"上传文件不能超过{maxSize}M");
                    }
                    if (!System.IO.Directory.Exists(path))
                    {
                        if (!string.IsNullOrEmpty(path))
                        {
                            //不存在路径则创建
                            System.IO.Directory.CreateDirectory(path);
                        }
                    };
                    //文件后缀名
                    string fileExt = System.IO.Path.GetExtension(file.FileName);
                    //新文件名
                    string fileName = Guid.NewGuid() + DateTime.Now.ToString("yyyyMMddHHmmss") + fileExt;
                    //完整路径
                    string filePath = System.IO.Path.Combine(path, fileName);
                    //写入本地
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                        stream.Flush();
                    }
                }
                else
                {
                    throw new Exception("上传文件不能为空");
                }
            }
            catch (Exception ex)
            {
                while (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                }
                throw ex;
            }
        }
        /// <summary>
        /// 文件上传（异步）
        /// </summary>
        /// <param name="path">存放地址</param>
        /// <param name="files">所有文件</param>
        /// <param name="maxSize">最大上传大小（默认10M）</param>
        public static async Task UploadAsync(string path, IFormFileCollection files, int maxSize = 10)
        {
            try
            {
                if (files != null && files.Sum(x => x.Length) > 0)
                {
                    foreach (IFormFile file in files)
                    {
                        await UploadAsync(path, file, maxSize);
                    }
                }
                else
                {
                    throw new Exception("上传文件不能为空");
                }
            }
            catch (Exception ex)
            {
                while (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                }
                throw ex;
            }
        }
        /// <summary>
        ///  文件上传（异步）
        /// </summary>
        /// <param name="path">存放地址</param>
        /// <param name="file">单个文件</param>
        /// <param name="maxSize">最大上传大小（默认10M）</param>
        public static async Task UploadAsync(string path, IFormFile file, int maxSize = 10)
        {
            try
            {
                if (file != null && file.Length > 0)
                {
                    if (file.Length > 1024 * 1024 * maxSize)
                    {
                        throw new Exception($"上传文件不能超过{maxSize}M");
                    }
                    if (!System.IO.Directory.Exists(path))
                    {
                        if (!string.IsNullOrEmpty(path))
                        {
                            //不存在路径则创建
                            System.IO.Directory.CreateDirectory(path);
                        }
                    };
                    //文件后缀名
                    string fileExt = System.IO.Path.GetExtension(file.FileName);
                    //新文件名
                    string fileName = Guid.NewGuid() + DateTime.Now.ToString("yyyyMMddHHmmss") + fileExt;
                    //完整路径
                    string filePath = System.IO.Path.Combine(path, fileName);
                    //写入本地
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                        await stream.FlushAsync();
                    }
                }
                else
                {
                    throw new Exception("上传文件不能为空");
                }
            }
            catch (Exception ex)
            {
                while (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                }
                throw ex;
            }
        }
        /// <summary>
        /// 获取文件contentType
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>contentType</returns>
        public static string GetContentType(string filePath)
        {
            string fileExt = System.IO.Path.GetExtension(filePath);
            var provider = new FileExtensionContentTypeProvider();
            string contentType = provider.Mappings[fileExt];
            return contentType;
        }
    }
}
