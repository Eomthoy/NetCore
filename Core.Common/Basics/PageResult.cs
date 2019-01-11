using System;
using System.Collections.Generic;
using System.Linq;

namespace Core.Common.Basics
{
    /// <summary>
    /// 数据分页
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class PageResult<T>
    {
        /// <summary>
        /// 数据总条数
        /// </summary>
        public int Total { get; set; }
        /// <summary>
        /// 页码
        /// </summary>
        public int Page { get; set; }
        /// <summary>
        /// 页码大小
        /// </summary>
        public int PageSize { get; set; }
        /// <summary>
        /// 总页数
        /// </summary>
        public int? TotalPages { get; set; }
        /// <summary>
        /// 数据
        /// </summary>
        public IEnumerable<T> Rows { get; set; }
        public PageResult()
        {
        }
        /// <summary>
        /// 生成分页数据
        /// </summary>
        /// <param name="page">页码</param>
        /// <param name="pageSize">页码大小</param>
        /// <param name="data">数据集</param>
        public PageResult(int page, int pageSize, IEnumerable<T> data)
        {
            this.Total = data.Count();
            this.Page = page;
            this.PageSize = pageSize;
            this.TotalPages = Convert.ToInt32(Math.Ceiling(data.Count() * 1.0 / pageSize));
            this.Rows = data.Skip((page - 1) * PageSize).Take(PageSize);
        }
    }
}
