package com.tools;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Random;

/**
 * Excel测试数据生成器
 * 用于生成大量测试数据以测试Excel处理工具的性能
 *
 * 使用方式：
 * java -cp target/excel-processor-1.0-SNAPSHOT-jar-with-dependencies.jar com.tools.TestDataGenerator [目标大小MB]
 *
 * 示例：
 * java -cp target/excel-processor-1.0-SNAPSHOT-jar-with-dependencies.jar com.tools.TestDataGenerator 150
 */
public class TestDataGenerator {

    // ==================== 配置常量 ====================
    private static final int DEFAULT_TARGET_SIZE_MB = 30;
    private static final int MIN_TARGET_SIZE_MB = 140;  // 允许的最小值，目标值-10MB
    private static final int MEMORY_WINDOW = 100;  // SXSSF内存窗口大小
    private static final int PROGRESS_INTERVAL = 10000;  // 进度输出间隔
    private static final int SIZE_CHECK_INTERVAL = 50000;  // 文件大小检查间隔

    // ==================== 真实姓氏库（100个常见姓氏）====================
    private static final String[] SURNAMES = {
        "王", "李", "张", "刘", "陈", "杨", "黄", "赵", "吴", "周",
        "徐", "孙", "马", "朱", "胡", "郭", "何", "高", "林", "罗",
        "郑", "梁", "谢", "宋", "唐", "许", "韩", "冯", "邓", "曹",
        "彭", "曾", "肖", "田", "董", "袁", "潘", "于", "蒋", "蔡",
        "余", "杜", "叶", "程", "苏", "魏", "吕", "丁", "任", "沈",
        "姚", "卢", "姜", "崔", "钟", "谭", "陆", "汪", "范", "金",
        "石", "廖", "贾", "夏", "韦", "付", "方", "白", "邹", "孟",
        "熊", "秦", "邱", "江", "尹", "薛", "闫", "段", "雷", "侯",
        "龙", "史", "陶", "黎", "贺", "顾", "毛", "郝", "龚", "邵",
        "万", "钱", "严", "覃", "武", "戴", "莫", "孔", "向", "汤"
    };

    // ==================== 真实名字库（200个常用字）====================
    private static final String[] GIVEN_NAMES = {
        "伟", "芳", "娜", "秀英", "敏", "静", "丽", "强", "磊", "军",
        "洋", "勇", "艳", "杰", "娟", "涛", "明", "超", "秀兰", "霞",
        "平", "刚", "桂英", "建华", "文", "华", "金凤", "素英", "玲", "建国",
        "慧", "峰", "秀珍", "雪", "辉", "莉", "萍", "燕", "云", "丹",
        "红", "梅", "林", "玉兰", "鑫", "浩", "琴", "晶", "颖", "成",
        "帆", "婷", "俊", "凯", "佳", "龙", "鹏", "阳", "宇", "欣",
        "琳", "斌", "亮", "畅", "楠", "想", "雨", "志", "博", "祥",
        "悦", "维", "瑞", "越", "腾", "睿", "远", "航", "坤", "健",
        "豪", "晨", "妍", "轩", "聪", "诗", "锐", "晓", "梓", "杰",
        "韵", "睿", "涵", "琪", "宁", "毅", "洁", "丽", "玉", "敏",
        "群", "海", "兰", "英", "秀", "青", "春", "波", "冰", "雅",
        "宏", "威", "东", "锋", "彬", "智", "凡", "安", "思", "飞",
        "宝", "天", "生", "冬", "兴", "光", "良", "园", "钢", "新",
        "星", "学", "康", "清", "庆", "永", "友", "松", "泉", "义",
        "元", "炜", "琦", "娴", "翔", "莹", "苗", "振", "珍", "达",
        "卫", "翠", "福", "旭", "彦", "虎", "运", "贵", "德", "斌",
        "宸", "瑶", "琰", "源", "森", "雯", "蕾", "瑜", "昊", "哲",
        "彤", "然", "珊", "恒", "岩", "炫", "城", "钰", "曦", "晗",
        "灿", "榕", "璇", "曼", "澄", "泽", "承", "恩", "湛", "霖",
        "煜", "昕", "墨", "璐", "皓", "烨", "熠", "翰", "泓", "澜"
    };

    // ==================== 客户备注模板（30种）====================
    private static final String[] CUSTOMER_REMARKS = {
        "意向客户，待进一步跟进", "已成交，满意度较高", "潜在客户，需要培养",
        "老客户，忠诚度高", "首次接触，了解产品中", "已签约，等待交付",
        "通过朋友推荐而来", "对价格比较敏感", "预算充足，决策快",
        "犹豫不决，需要多次沟通", "竞品用户，考虑转换", "VIP客户，重点维护",
        "投诉客户，需要安抚", "已流失，尝试挽回", "合作意向强，正在谈判",
        "试用期客户，观察中", "高价值客户，优先服务", "决策周期长，耐心跟进",
        "有明确需求，匹配度高", "预算有限，推荐基础版", "企业客户，采购流程复杂",
        "个人客户，需求简单", "行业标杆客户，影响力大", "合同即将到期，续约跟进",
        "新注册用户，首次咨询", "多次咨询，未下单", "对比多家供应商",
        "需要定制化方案", "转介绍客户，信任度高", "咨询后失联，待激活"
    };

    // ==================== 其他备注模板（20种长文本）====================
    private static final String[] OTHER_REMARKS = {
        "客户是通过线上广告了解到我们的产品，初步沟通后表示对功能比较感兴趣，但需要和团队内部讨论后再做决定。客户公司规模约50人，属于成长型企业，预算在10-20万之间。计划下周安排产品演示，重点展示我们在数据分析和报表生成方面的优势功能。",

        "该客户是老客户介绍过来的，信任度较高。目前使用竞品方案已有2年，对现有方案存在诸多不满，主要集中在售后服务和产品稳定性方面。我方优势明显，成交概率大。已安排技术团队准备详细方案，计划本周五上门演示。客户对价格不太敏感，更看重产品质量和服务。",

        "客户所在行业是医疗健康，对数据安全和隐私保护要求极高。需要提供相关资质证明和合规文档。客户决策周期较长，预计需要2-3个月。已提交初步方案，等待客户反馈。需要注意的是，客户采购流程较为严格，需要经过技术部、采购部、法务部三重审核。",

        "通过行业展会认识的客户，现场体验产品后印象深刻。客户企业规模较大，员工超过500人，是行业内知名企业。如能成交将对我们品牌影响力有很大提升。客户提出需要支持与现有ERP系统对接，技术部已确认可行。预计下月提交完整技术方案和报价。",

        "该客户之前使用过我们的产品，因公司业务调整暂停了服务。现在业务恢复，重新联系我们希望继续合作。对产品功能已经很熟悉，主要关注新版本的升级内容和价格。已安排客户经理专门对接，提供老客户优惠政策，预计本月内可以完成续约。",

        "客户是创业公司，团队年轻有活力，对新技术接受度高。虽然目前规模不大，但发展势头良好，有较大增长潜力。客户预算有限，建议推荐性价比高的套餐。已经过两轮沟通，客户对产品很满意，主要在等融资到账后签约，预计下月可以成交。",

        "大型企业客户，采购流程非常规范和复杂。需要参与招投标流程，准备详细的技术方案、资质文件、报价文件等。竞争对手较多，需要突出我们的差异化优势。客户特别关注实施周期和培训服务，这是我们的强项。已组建专项团队跟进此项目。",

        "客户通过搜索引擎找到我们，属于主动寻找解决方案的类型，购买意愿较强。经过初步沟通，客户需求与我们产品功能高度匹配。客户决策快，希望尽快上线使用。已安排明天进行线上产品演示，演示后预计可以快速推进合同签订。客户公司在上海，后续可能需要安排上门服务。",

        "行业KOL推荐的客户，对我们品牌有一定认知。客户比较理性，会详细对比多家产品的功能和价格。我们需要准备详细的产品对比资料，突出我们的核心优势。客户预计一个月内做出决策，目前处于信息收集阶段。已发送详细产品资料和案例介绍。",

        "客户之前咨询过但没有成交，这次是第二次联系。上次未成交的原因是预算不足，现在公司获得新一轮融资，预算充足。客户对产品已经比较了解，这次沟通主要确认新需求和最新报价。成交概率很高，需要抓紧跟进，避免被竞品抢走。",

        "VIP客户，使用我们产品已经3年，续费稳定，满意度高。这次联系是希望增购更多功能模块，扩大使用范围。客户对我们的产品和服务高度认可，经常推荐给同行。需要给予一定的老客户优惠和增值服务。已安排客户成功团队提供专属服务。",

        "客户在试用我们的免费版本，使用两周后觉得不错，咨询付费版本的功能和价格。试用期间客户提出了一些改进建议，产品团队已记录并计划在下个版本优化。客户对产品整体满意，主要考虑的是投入产出比。需要帮助客户算清楚ROI，促进转化。",

        "集团型客户，有多个分支机构，采购需求量大。目前先从一个分公司试点，效果好的话会推广到整个集团。这是一个长期的战略客户，需要投入足够的资源和精力。已安排高级售前顾问全程支持，确保试点项目成功实施，为后续扩展打好基础。",

        "政府机关客户，采购流程严格，需要走公开招标程序。我们需要准备完整的投标文件，包括公司资质、产品介绍、技术方案、报价等。政府项目周期长但稳定性好，一旦中标后续服务期可能长达3-5年。目前在准备标书阶段，预计下月开标。",

        "教育行业客户，学校采购预算有限但需求明确。客户比较关注教育行业的成功案例和特色功能。我们在教育行业有丰富经验，已经服务了多所知名高校。可以安排客户参观已有客户的使用情况。教育客户一般在年底有预算，需要把握时机。",

        "跨国企业中国区负责人，需要采购全球统一的解决方案。除了中国区使用，还需要考虑与国际总部系统的对接。客户对数据安全、多语言支持、全球化部署有特殊要求。这是一个复杂度较高的项目，需要技术团队深度参与方案设计。",

        "互联网公司客户，业务发展迅速，对产品的扩展性和性能要求高。客户技术团队实力强，会深入评估产品的技术架构。我们需要安排技术专家与客户的技术团队对接，解答各种技术细节问题。客户决策快，一旦认可技术方案会迅速推进采购。",

        "传统制造业客户，正在进行数字化转型，首次采购此类产品。客户对产品理解有限，需要详细的培训和引导。我们的机会在于可以成为客户数字化转型的重要合作伙伴。需要耐心沟通，帮助客户理解产品价值，同时提供完善的实施和培训服务。",

        "客户是行业内的龙头企业，品牌影响力大，如能成为我们的客户将极大提升我们的市场地位。客户对供应商的要求很高，会进行全方位的评估。我们需要展现出专业性和可靠性，获得客户信任。这是一个具有战略意义的客户，公司高层高度重视。",

        "价格敏感型客户，多次询价比价，总希望获得最低价格。客户虽然预算有限，但需求是真实的。建议推荐基础版本或提供分期付款方案，降低客户的采购门槛。需要注意控制成本，避免过度让利。同时要向客户传递产品的价值，而不仅仅是拼价格。"
    };

    // ==================== 跟进记录操作类型（10种）====================
    private static final String[] FOLLOW_UP_TYPES = {
        "电话沟通", "上门拜访", "产品演示", "方案讲解", "合同谈判",
        "售后回访", "问题处理", "邮件往来", "微信沟通", "会议讨论"
    };

    // ==================== 跟进记录内容模板（30种）====================
    private static final String[] FOLLOW_UP_CONTENTS = {
        "客户表示对产品核心功能非常满意，但希望能在价格上给予一定优惠。已向上级申请特批权限，预计明天可以给出最终报价。客户承诺如果价格合适，本周内可以签约。需要注意的是，客户有两家备选供应商，我们需要尽快确定方案。",

        "实地考察了客户公司，了解了他们的业务流程和痛点。发现客户当前使用的系统存在严重的效率问题，我们的方案可以完美解决。客户负责人对演示效果非常认可，要求尽快出具详细方案。现场还见到了客户的技术总监，初步沟通了技术对接事宜。",

        "客户提出了一些定制化需求，主要涉及报表功能和权限管理。已记录详细需求，转交技术部门评估。预计3个工作日内可以给出技术方案和报价。客户表示愿意等待，并强调定制功能对他们业务很重要，是决定是否采购的关键因素。",

        "与客户进行了深入的方案讨论，客户对我们提出的整体解决方案表示认可。会议持续了2个小时，客户方有技术总监、运营总监、财务总监参加。重点讨论了系统集成、数据迁移、人员培训等实施细节。客户要求我们提供详细的项目实施计划。",

        "客户反馈在使用过程中遇到一些问题，主要是操作不熟练导致的。通过远程协助，帮助客户解决了所有问题，并提供了详细的操作指南。客户对我们的响应速度和服务态度表示满意。这次良好的服务体验增强了客户对我们的信任。",

        "合同谈判进入关键阶段，双方就价格、付款方式、服务条款等进行了细致讨论。我方在价格上做了适当让步，客户在付款方式上也做了妥协，最终达成一致。预计下周可以正式签署合同。法务部门正在审核最终版本的合同条款。",

        "向客户演示了最新版本的产品功能，客户对新增的数据分析模块特别感兴趣。演示过程很顺利，客户频频点头表示认可。演示后客户提出希望能试用一个月，已安排技术支持团队协助客户开通试用账号。试用期间需要密切跟进客户的使用情况。",

        "通过微信与客户保持日常联系，了解客户最新的业务动态。客户提到公司最近在拓展新业务，可能会有新的系统需求。已记录客户的新需求，后续安排专门的需求沟通会议。保持与客户的良好关系对长期合作很重要。",

        "客户公司进行了组织架构调整，我们的对接人发生了变更。主动联系了新的负责人，介绍了我们的产品和服务，重新建立联系。新负责人表示会先了解现有系统的使用情况，之后再讨论是否需要升级或扩展。需要给新负责人一些时间熟悉业务。",

        "参加了客户组织的供应商大会，与客户高层进行了面对面交流。会上了解到客户未来的发展规划和采购需求，这为我们制定销售策略提供了重要参考。客户对我们过去一年的服务表示满意，希望继续深化合作关系。",

        "客户反馈系统运行稳定，使用效果良好，已经成为公司日常运营的重要工具。客户提出希望能增加一些高级功能，已记录需求并报价。这是一个增购的好机会，需要抓紧跟进。客户表示下季度有预算，可以考虑升级到高级版本。",

        "处理了客户的一个紧急问题，客户系统突然无法登录，影响了业务运营。技术团队快速响应，15分钟内定位问题并修复。问题是由于客户网络配置变更导致的。客户对我们的应急响应能力表示高度认可，这次事件反而增强了客户对我们的信任。",

        "发送了详细的产品对比资料给客户，帮助客户了解我们与竞品的差异。重点突出了我们在用户体验、技术创新、服务质量方面的优势。客户回复表示资料很有帮助，目前正在内部评审，预计两周内会给出初步决策意见。",

        "客户邀请我们参加他们的内部评审会议，向他们的管理团队介绍我们的解决方案。准备了精美的PPT和演示视频，会议进行得很成功。客户管理团队对方案给予了积极评价，要求采购部门尽快推进合同签署流程。这是一个重要的里程碑。",

        "跟进客户的试用情况，了解到客户在使用过程中发现了一些小bug。已记录问题并提交给开发团队，承诺在下个版本中修复。虽然有一些小问题，但客户对产品的整体评价还是正面的，表示试用结束后会考虑正式采购。",

        "客户合同即将到期，主动联系客户讨论续约事宜。客户对过去一年的合作表示满意，愿意继续合作。讨论了新的合作周期和优惠方案，客户要求能在原价基础上给予一定折扣。已向上级汇报，申请续约优惠政策。",

        "举办了客户答谢会，邀请了包括该客户在内的多家重要客户参加。活动现场气氛热烈，客户之间也进行了交流和分享。这位客户在会上分享了使用我们产品的成功经验，对我们的品牌传播起到了很好的作用。活动后客户关系更加紧密。",

        "收到客户的投诉，反映产品某个功能使用不便。立即安排产品经理与客户沟通，了解详细的使用场景和改进建议。产品经理表示会在下个迭代中优化该功能。客户对我们重视用户反馈的态度表示认可，投诉得到了妥善处理。",

        "为客户提供了一次免费的产品培训，帮助客户团队更好地使用系统的高级功能。培训持续了半天，客户参训人员约20人。培训效果很好，客户表示很多之前不知道的功能原来如此强大。培训后客户的产品使用深度明显提升。",

        "客户推荐了一个新客户给我们，这是对我们服务最好的肯定。已经联系了新客户，初步沟通很顺利。需要向原客户表示感谢，考虑提供一些增值服务作为回馈。客户转介绍是最优质的销售线索，成交率往往很高。",

        "与客户讨论了未来的合作规划，客户提出希望能在产品功能定制方面进行更深入的合作。这超出了标准产品销售的范畴，需要公司高层决策是否接受定制开发项目。已将客户需求整理成报告提交给管理层，等待进一步指示。",

        "客户参加了我们举办的行业研讨会，会上展示了最新的产品发展路线图。客户对我们的技术创新能力和行业洞察力表示赞赏。会后客户主动询问新功能的上线时间，表示愿意成为新功能的首批用户。这是一个深化合作的好机会。",

        "处理客户的账单问题，客户对某项收费存在疑问。详细核对了合同条款，向客户解释了收费的依据。客户理解后表示接受，并为之前的误解表示歉意。这次沟通虽然是为了解决问题，但也是一次增进相互理解的机会。",

        "客户提出想要参观我们公司，了解我们的团队和研发实力。精心安排了参观路线，介绍了公司文化、团队构成、研发流程等。客户参观后对我们的专业性和规范性表示高度认可，信心进一步增强。这次参观对促进合作起到了积极作用。",

        "收到客户的表扬信，感谢我们在项目实施过程中的专业服务和敬业精神。这封表扬信被公司用作内部激励的案例，也可以作为对外宣传的素材。客户的认可是对团队最大的鼓励，需要继续保持高质量的服务水平。",

        "与客户进行季度业务回顾，总结了过去三个月的合作情况，讨论了存在的问题和改进方向。客户提出了一些中肯的建议，我们虚心接受并承诺改进。定期的业务回顾有助于保持良好的客户关系，及时发现和解决问题。",

        "客户公司进行了系统升级，需要我们的产品也相应升级以保持兼容。技术团队评估后确认可以升级，并给出了升级方案和时间表。客户对我们的快速响应表示满意，双方约定了具体的升级时间，确保不影响客户的正常业务。",

        "应客户要求提供了详细的ROI分析报告，帮助客户向上级申请预算。报告中详细说明了使用我们产品后可以带来的效率提升和成本节约。客户表示报告对他们很有帮助，将在下周的预算会议上使用。这种专业的销售支持材料很重要。",

        "客户反馈竞争对手主动联系他们，开出了很有吸引力的条件。及时与客户沟通，了解竞品的具体方案和报价。针对竞品的优势，强调我们的差异化价值。客户表示会综合考虑，不会轻易更换供应商。需要密切关注竞争动态，制定应对策略。",

        "组织了一次客户满意度调研，收集客户对产品和服务的评价。该客户给出了较高的满意度评分，同时也提出了一些改进建议。这些宝贵的反馈将帮助我们持续优化产品和服务。已将调研结果整理成报告，提交给产品和服务团队参考。"
    };

    private static final Random random = new Random();

    public static void main(String[] args) {
        // 解析命令行参数
        int targetSizeMB = DEFAULT_TARGET_SIZE_MB;
        if (args.length > 0) {
            try {
                targetSizeMB = Integer.parseInt(args[0]);
                if (targetSizeMB < 10) {
                    System.out.println("目标大小太小，最小值为10MB，已自动调整为10MB");
                    targetSizeMB = 10;
                } else if (targetSizeMB > 1000) {
                    System.out.println("目标大小太大，最大值为1000MB，已自动调整为1000MB");
                    targetSizeMB = 1000;
                }
            } catch (NumberFormatException e) {
                System.out.println("参数格式错误，使用默认值: " + DEFAULT_TARGET_SIZE_MB + "MB");
            }
        }

        // 输出启动信息
        printHeader(targetSizeMB);

        // 生成文件名（带时间戳）
        String fileName = generateFileName();
        File outputFile = new File(fileName);

        System.out.println("输出文件: " + outputFile.getAbsolutePath());
        System.out.println("开始生成...\n");

        long startTime = System.currentTimeMillis();

        try {
            generateExcelFile(outputFile, targetSizeMB);

            long endTime = System.currentTimeMillis();
            long duration = (endTime - startTime) / 1000;

            printFooter(outputFile, duration);

        } catch (Exception e) {
            System.err.println("\n生成失败: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 生成Excel文件
     */
    private static void generateExcelFile(File outputFile, int targetSizeMB) throws Exception {
        SXSSFWorkbook workbook = new SXSSFWorkbook(MEMORY_WINDOW);

        try {
            Sheet sheet = workbook.createSheet("客户数据");

            // 创建表头样式
            CellStyle headerStyle = createHeaderStyle(workbook);

            // 写入表头
            Row headerRow = sheet.createRow(0);
            String[] headers = {"客户姓名", "客户电话", "客户备注", "客户星级", "其他备注", "跟进记录"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            // 设置列宽
            sheet.setColumnWidth(0, 15 * 256);  // 客户姓名
            sheet.setColumnWidth(1, 15 * 256);  // 客户电话
            sheet.setColumnWidth(2, 25 * 256);  // 客户备注
            sheet.setColumnWidth(3, 12 * 256);  // 客户星级
            sheet.setColumnWidth(4, 60 * 256);  // 其他备注
            sheet.setColumnWidth(5, 80 * 256);  // 跟进记录

            // 生成数据行
            // 策略：根据目标大小估算需要生成的行数
            // 经验值：每MB约需要5200行（基于XLSX压缩）
            // 150MB需要约780,000行
            int estimatedRows = targetSizeMB * 5200;
            int rowNum = 1;

            System.out.println("预估需要生成约 " + String.format("%,d", estimatedRows) + " 行数据\n");

            while (rowNum <= estimatedRows) {
                Row row = sheet.createRow(rowNum);

                // 生成各列数据
                row.createCell(0).setCellValue(generateName());
                row.createCell(1).setCellValue(generatePhone());
                row.createCell(2).setCellValue(generateCustomerRemark());
                row.createCell(3).setCellValue(generateStarLevel());
                row.createCell(4).setCellValue(generateOtherRemark());
                row.createCell(5).setCellValue(generateFollowUpRecords());

                rowNum++;

                // 定期输出进度
                if (rowNum % PROGRESS_INTERVAL == 0) {
                    int progress = (int) ((double) rowNum / estimatedRows * 100);
                    System.out.printf("已生成: %,d 行 (进度: %d%%)\n", rowNum, progress);
                }
            }

            System.out.printf("\n已生成 %,d 行数据，正在保存文件...\n", rowNum - 1);

            // 最终保存
            try (FileOutputStream out = new FileOutputStream(outputFile)) {
                workbook.write(out);
            }

            System.out.println("文件保存完成！\n");

        } finally {
            workbook.close();
            workbook.dispose();
        }
    }

    /**
     * 创建表头样式
     */
    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * 生成随机客户姓名
     */
    private static String generateName() {
        String surname = SURNAMES[random.nextInt(SURNAMES.length)];
        String givenName;

        // 70%概率生成2字名，30%概率生成1字名
        if (random.nextInt(100) < 70) {
            givenName = GIVEN_NAMES[random.nextInt(GIVEN_NAMES.length)];
        } else {
            // 1字名，从GIVEN_NAMES中选择单字
            String fullName = GIVEN_NAMES[random.nextInt(GIVEN_NAMES.length)];
            givenName = fullName.length() > 1 && random.nextBoolean() ?
                fullName.substring(0, 1) : fullName;
        }

        return surname + givenName;
    }

    /**
     * 生成随机手机号码
     */
    private static String generatePhone() {
        String[] prefixes = {"130", "131", "132", "133", "134", "135", "136", "137", "138", "139",
                            "150", "151", "152", "153", "155", "156", "157", "158", "159",
                            "180", "181", "182", "183", "184", "185", "186", "187", "188", "189",
                            "170", "171", "175", "176", "177", "178"};

        String prefix = prefixes[random.nextInt(prefixes.length)];
        StringBuilder phone = new StringBuilder(prefix);

        for (int i = 0; i < 8; i++) {
            phone.append(random.nextInt(10));
        }

        return phone.toString();
    }

    /**
     * 生成随机客户备注
     */
    private static String generateCustomerRemark() {
        return CUSTOMER_REMARKS[random.nextInt(CUSTOMER_REMARKS.length)];
    }

    /**
     * 生成随机客户星级
     */
    private static String generateStarLevel() {
        // 权重分布: 5星(15%), 4星(25%), 3星(35%), 2星(20%), 1星(5%)
        int rand = random.nextInt(100);
        int stars;

        if (rand < 5) {
            stars = 1;
        } else if (rand < 25) {
            stars = 2;
        } else if (rand < 60) {
            stars = 3;
        } else if (rand < 85) {
            stars = 4;
        } else {
            stars = 5;
        }

        // 50%概率返回星号，50%返回数字+文字
        if (random.nextBoolean()) {
            return repeatString("★", stars) + repeatString("☆", 5 - stars);
        } else {
            return stars + "星";
        }
    }

    /**
     * 重复字符串（Java 8兼容方法）
     */
    private static String repeatString(String str, int count) {
        if (count <= 0) {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < count; i++) {
            sb.append(str);
        }
        return sb.toString();
    }

    /**
     * 生成随机其他备注（长文本）
     */
    private static String generateOtherRemark() {
        // 80%概率生成长备注，20%概率生成短备注
        if (random.nextInt(100) < 80) {
            return OTHER_REMARKS[random.nextInt(OTHER_REMARKS.length)];
        } else {
            // 短备注：随机组合2-3个客户备注
            int count = random.nextInt(2) + 2;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < count; i++) {
                if (i > 0) sb.append("。");
                sb.append(CUSTOMER_REMARKS[random.nextInt(CUSTOMER_REMARKS.length)]);
            }
            return sb.toString();
        }
    }

    /**
     * 生成随机跟进记录（多条记录，换行分隔）
     */
    private static String generateFollowUpRecords() {
        // 生成3-5条跟进记录
        int recordCount = random.nextInt(3) + 3;
        StringBuilder sb = new StringBuilder();

        for (int i = 0; i < recordCount; i++) {
            if (i > 0) {
                sb.append("\n\n");
            }

            // 生成日期（最近90天内的随机日期）
            int daysAgo = random.nextInt(90);
            String date = generateDateString(daysAgo);

            // 随机选择操作类型
            String type = FOLLOW_UP_TYPES[random.nextInt(FOLLOW_UP_TYPES.length)];

            // 随机选择内容
            String content = FOLLOW_UP_CONTENTS[random.nextInt(FOLLOW_UP_CONTENTS.length)];

            sb.append(String.format("[%s] %s：%s", date, type, content));
        }

        return sb.toString();
    }

    /**
     * 生成日期字符串（当前日期往前推N天）
     */
    private static String generateDateString(int daysAgo) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        long timestamp = System.currentTimeMillis() - (long) daysAgo * 24 * 60 * 60 * 1000;
        return sdf.format(new Date(timestamp));
    }

    /**
     * 生成带时间戳的文件名
     */
    private static String generateFileName() {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String timestamp = sdf.format(new Date());
        return "test_data_" + timestamp + ".xlsx";
    }

    /**
     * 获取文件大小（MB）
     */
    private static long getFileSizeMB(File file) {
        if (!file.exists()) {
            return 0;
        }
        return file.length() / (1024 * 1024);
    }

    /**
     * 打印启动信息
     */
    private static void printHeader(int targetSizeMB) {
        System.out.println("========================================");
        System.out.println("     Excel测试数据生成器");
        System.out.println("========================================");
        System.out.println("目标文件大小: " + targetSizeMB + " MB");
    }

    /**
     * 打印完成信息
     */
    private static void printFooter(File outputFile, long durationSeconds) {
        long fileSizeMB = getFileSizeMB(outputFile);

        System.out.println("\n========================================");
        System.out.println("          生成完成！");
        System.out.println("========================================");
        System.out.println("输出文件: " + outputFile.getAbsolutePath());

        // 计算行数（需要重新读取，但为了避免内存问题，这里估算）
        // 实际行数会在文件生成过程中记录
        System.out.println("文件大小: " + fileSizeMB + " MB");

        long minutes = durationSeconds / 60;
        long seconds = durationSeconds % 60;
        System.out.println("耗时: " + minutes + "分" + seconds + "秒");
        System.out.println("========================================");
    }
}
