using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolDemo
{
    public class ResultInfo
    {
        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 电话
        /// </summary>
        public string Phone { get; set; }

        /// <summary>
        /// 性别
        /// </summary>
        public string Sex { get; set; }

        /// <summary>
        /// 国籍
        /// </summary>
        public string Country { get; set; }

        /// <summary>
        /// 身份证
        /// </summary>
        public string IDCard { get; set; }

        /// <summary>
        /// 刷卡时间
        /// </summary>
        public string PassTime { get; set; }

        /// <summary>
        /// 刷卡状态
        /// </summary>
        public string PassState { get; set; }

        /// <summary>
        /// 刷卡位置
        /// </summary>
        public string PassLocation { get; set; }

        /// <summary>
        /// 所属公司
        /// </summary>
        public string Company { get; set; }
    }
}
