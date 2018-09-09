using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace NPOIExtension.Web.Models
{
    public class SettlementModel
    {
        /// <summary>
        /// 主键
        /// </summary>
        [DisplayName("主键")]
        public int Id { get; set; }

        /// <summary>
        /// 业务类型
        /// </summary>
        [DisplayName("业务类型")]
        public int BussinessType { get; set; }

        /// <summary>
        /// 对账单流水号
        /// </summary>  
        [DisplayName("对账单流水号")]
        public string SerialId { get; set; }
        /// <summary>
        /// 对账单标题
        /// </summary>
        [DisplayName("对账单标题")]
        public string Title { get; set; }

        /// <summary>
        /// 生单开始时间
        /// </summary>
        [DisplayName("生单开始时间")]
        public DateTime CreateDateBegin { get; set; }

        /// <summary>
        /// 生单结束时间
        /// </summary>
        [DisplayName("生单结束时间")]
        public DateTime CreateDateEnd { get; set; }

    }
}