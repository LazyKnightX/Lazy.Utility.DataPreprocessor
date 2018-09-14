using System.Collections.Generic;

namespace Lazy.Utility.DataPreprocessor.Data.RPGMakerMV
{
  #region Objects

  public class RMObjectBase
  {
    /// <summary>
    /// RM: 左侧ID
    /// EX: ID
    /// </summary>
    public int id = 0;

    /// <summary>
    /// RM: 备注
    /// </summary>
    public string note = "";

    /// <summary>
    /// <para>RM: 名称</para>
    /// <para>EX: 名称</para>
    /// </summary>
    public string name = "";
  }

  public class Animation
  {
    public int id = 0;
    public int animation1Hue { get; set; }
    public string animation1Name { get; set; }
    public int animation2Hue { get; set; }
    public string animation2Name { get; set; }
    public List<List<List<int>>> frames { get; set; }
    public string name = "";
    public int position { get; set; }
    public List<Timing> timings { get; set; }
  }

  public class CommonEvent
  {
    public int id { get; set; }
    public List<ActionList> list { get; set; }
    public string name { get; set; }
    public int switchId { get; set; }
    public int trigger { get; set; }
  }

  public class Item : RMObjectBase
  {
    /// <summary>
    /// RM: 动画
    /// EX: 动画
    /// </summary>
    public int animationId { get; set; }

    /// <summary>
    /// RM: 消耗品
    /// EX: 消耗品
    /// </summary>
    public bool consumable { get; set; }

    /// <summary>
    /// RM: 伤害
    /// EX: ...复合数据
    /// </summary>
    public Damage damage { get; set; }

    /// <summary>
    /// RM: 说明
    /// EX: 描述
    /// </summary>
    public string description { get; set; }

    /// <summary>
    /// RM: 效果
    /// </summary>
    public List<Effect> effects { get; set; }

    /// <summary>
    /// RM: 命中类型(0: 必定命中, 1: 物理攻击, 2: 魔法攻击)
    /// EX: 命中类型
    /// </summary>
    public int hitType { get; set; }

    /// <summary>
    /// RM: 图标
    /// </summary>
    public int iconIndex { get; set; }

    /// <summary>
    /// <para>RM: 物品类型</para>
    /// <para>EX: 物品类型</para>
    /// <para>0: 普通物品, 1: 重要物品, 2: 隐藏物品 A, 3: 隐藏物品 B</para>
    /// </summary>
    public int itypeId { get; set; }

    /// <summary>
    /// RM: 使用场合(0: 随时可用, 1: 战斗画面, 2: 菜单画面, 3: 不能使用)
    /// EX: 场合
    /// </summary>
    public int occasion { get; set; }

    /// <summary>
    /// RM: 价格
    /// EX: 价格
    /// </summary>
    public int price { get; set; }

    /// <summary>
    /// RM: 连续次数 (1~9)
    /// EX: 连续
    /// </summary>
    public int repeats { get; set; }

    /// <summary>
    /// RM: 范围(0: 无, 1: 敌方单体, 2: 敌方全体, 3: 随机1个敌人, 4: 随机2个敌人, 5: 随机3个敌人, 6: 随机4个敌人, 7: 我方单体, 8: 我方全体, 9: 我方单体(无法战斗), 10: 我方全体(无法战斗), 11: 使用者)
    /// EX: 对象(物品表)
    /// </summary>
    public int scope { get; set; }

    /// <summary>
    /// RM: 速度补正
    /// EX: 发动速度
    /// </summary>
    public int speed { get; set; }

    /// <summary>
    /// RM: 成功率(100% = 100)
    /// EX: 成功率
    /// </summary>
    public int successRate { get; set; }

    /// <summary>
    /// RM: 获得TP
    /// </summary>
    public int tpGain { get; set; }
  }

  public class Weapon : RMObjectBase
  {
    public int animationId { get; set; }
    public string description { get; set; }
    public int etypeId { get; set; }
    public List<Trait> traits { get; set; }
    public int iconIndex { get; set; }
    public List<int> @params { get; set; }
    public int price { get; set; }
    public int wtypeId { get; set; }
  }

  public class Skill : RMObjectBase
  {
    public int animationId { get; set; }
    public Damage damage { get; set; }
    public string description { get; set; }
    public List<Effect> effects { get; set; }
    public int hitType { get; set; }
    public int iconIndex { get; set; }
    public string message1 { get; set; }
    public string message2 { get; set; }
    public int mpCost { get; set; }
    public int occasion { get; set; }
    public int repeats { get; set; }
    public int requiredWtypeId1 { get; set; }
    public int requiredWtypeId2 { get; set; }

    /// <summary>
    /// RM: 范围(0: 无, 1: 敌方单体, 2: 敌方全体, 3: 随机1个敌人, 4: 随机2个敌人, 5: 随机3个敌人, 6: 随机4个敌人, 7: 我方单体, 8: 我方全体, 9: 我方单体(无法战斗), 10: 我方全体(无法战斗), 11: 使用者)
    /// EX: 对象(物品表)
    /// </summary>
    public int scope { get; set; }

    /// <summary>
    /// RM: 速度补正(整数)
    /// EX: 发动速度
    /// </summary>
    public int speed { get; set; }

    public int stypeId { get; set; }

    /// <summary>
    /// RM: 成功率(100% = 100)
    /// EX: 成功率
    /// </summary>
    public int successRate { get; set; }

    public int tpCost { get; set; }
    public int tpGain { get; set; }
  }

  public class Enemy : RMObjectBase
  {
    public List<RMAction> actions { get; set; }
    public int battlerHue { get; set; }
    public string battlerName { get; set; }

    /// <summary>
    /// 掉落物品
    /// </summary>
    public List<DropItem> dropItems { get; set; }

    public int exp { get; set; }
    public List<Trait> traits { get; set; }
    public int gold { get; set; }
    public List<int> @params { get; set; }
  }

  public class Trader
  {
    public int id = 0;
    public string @interface = "";
    public string draw { get; set; }
    public string coin { get; set; }
    public string name { get; set; }
    public int level { get; set; }
    public List<string> chats { get; set; }
    public Goods sell { get; set; }
    public Goods buy { get; set; }
    public List<List<double>> discount { get; set; }
    public int discountCount { get; set; }
  }

  public class Armor : RMObjectBase
  {
    public int atypeId { get; set; }
    public string description { get; set; }
    public int etypeId { get; set; }
    public List<Trait> traits { get; set; }
    public int iconIndex { get; set; }
    public List<int> @params { get; set; }
    public int price { get; set; }
  }

  public class State : RMObjectBase
  {
    public int autoRemovalTiming { get; set; }
    public int chanceByDamage { get; set; }
    public List<Trait> traits { get; set; }
    public int iconIndex { get; set; }
    public int maxTurns { get; set; }
    public string message1 { get; set; }
    public string message2 { get; set; }
    public string message3 { get; set; }
    public string message4 { get; set; }
    public int minTurns { get; set; }
    public int motion { get; set; }
    public int overlay { get; set; }
    public int priority { get; set; }
    public bool removeAtBattleEnd { get; set; }
    public bool removeByDamage { get; set; }
    public bool removeByRestriction { get; set; }
    public bool removeByWalking { get; set; }
    public int restriction { get; set; }
    public int stepsToRemove { get; set; }
  }

  #endregion

  #region Database

  public class RMItem : List<Item>
  {
  }

  public class RMWeapon : List<Weapon>
  {
  }

  public class RMSkill : List<Skill>
  {
  }

  public class RMEnemy : List<Enemy>
  {
  }

  public class RMTrader : List<Trader>
  {
  }

  public class RMArmor : List<Armor>
  {
  }

  public class RMAnimation : List<Animation>
  {
  }

  public class RMState : List<State>
  {
  }

  public class RMCommonEvent : List<CommonEvent>
  {
  }

  public class RMSystem
  {
    public Airship airship { get; set; }
    public List<string> armorTypes { get; set; }
    public List<AttackMotion> attackMotions { get; set; }
    public Sound battleBgm { get; set; }
    public string battleback1Name { get; set; }
    public string battleback2Name { get; set; }
    public int battlerHue { get; set; }
    public string battlerName { get; set; }
    public Boat boat { get; set; }
    public string currencyUnit { get; set; }
    public Sound defeatMe { get; set; }
    public int editMapId { get; set; }

    /// <summary>
    /// RM: F9->类型->属性 (0: 无, 1+: 用户自定义)
    /// </summary>
    public List<string> elements { get; set; }

    public List<string> equipTypes { get; set; }
    public string gameTitle { get; set; }
    public Sound gameoverMe { get; set; }
    public string locale { get; set; }
    public List<object> magicSkills { get; set; }
    public List<bool> menuCommands { get; set; }
    public bool optDisplayTp { get; set; }
    public bool optDrawTitle { get; set; }
    public bool optExtraExp { get; set; }
    public bool optFloorDeath { get; set; }
    public bool optFollowers { get; set; }
    public bool optSideView { get; set; }
    public bool optSlipDeath { get; set; }
    public bool optTransparent { get; set; }
    public List<int> partyMembers { get; set; }
    public Ship ship { get; set; }
    public List<string> skillTypes { get; set; }
    public List<Sound> sounds { get; set; }
    public int startMapId { get; set; }
    public int startX { get; set; }
    public int startY { get; set; }
    public List<string> switches { get; set; }
    public Terms terms { get; set; }
    public List<TestBattler> testBattlers { get; set; }
    public int testTroopId { get; set; }
    public string title1Name { get; set; }
    public string title2Name { get; set; }
    public Sound titleBgm { get; set; }
    public List<string> variables { get; set; }
    public int versionId { get; set; }
    public Sound victoryMe { get; set; }
    public List<string> weaponTypes { get; set; }
    public List<int> windowTone { get; set; }
  }

  #endregion

  public class Damage
  {
    /// <summary>
    /// RM: 暴击
    /// EX: 暴击(物品表)
    /// </summary>
    public bool critical { get; set; }

    /// <summary>
    /// <para>RM: 属性 (-1: 普通攻击, 0: 无, 1+: 用户定义)</para>
    /// <para>EX: 属性(物品表)</para>
    /// <para>这是一个开发者自定义的字段。 (F9->类型->属性, $dataSystem.elements)</para>
    /// </summary>
    public int elementId { get; set; }

    /// <summary>
    /// RM: 计算公式
    /// EX: 计算公式(物品表)
    /// </summary>
    public string formula { get; set; }

    /// <summary>
    /// RM: 类型 (0: 无, 1: HP伤害, 2: MP伤害, 3: HP恢复, 4: MP恢复, 5: HP吸收, 6: MP吸收)
    /// EX: 伤害类型(物品表)
    /// </summary>
    public int type { get; set; }

    /// <summary>
    /// 伤害的随机补正范围，会在造成伤害的时候额外造成这个值转换为百分比后的伤害。
    /// RM: 分散度(100% = 100)
    /// EX: 离散度(物品表)
    /// </summary>
    public int variance { get; set; }
  }

  /// <summary>
  /// RM: 效果
  /// </summary>
  public class Effect
  {
    /// <summary>
    /// RM: 效果类型，如“解除状态”
    /// </summary>
    public int code { get; set; }

    /// <summary>
    /// RM: 内容，如“解除状态”的“无法战斗”、“剧毒”，根据不同的效果类型而有不同的值
    /// </summary>
    public int dataId { get; set; }

    /// <summary>
    /// RM: 参数1，根据不同的效果类型有不同的值
    /// </summary>
    public double value1 { get; set; }

    /// <summary>
    /// RM: 参数2，根据不同的效果类型有不同的值
    /// </summary>
    public double value2 { get; set; }
  }

  public class Trait
  {
    public int code { get; set; }
    public int dataId { get; set; }
    public double value { get; set; }
  }

  public class Sound
  {
    public string name { get; set; }
    public int pan { get; set; }
    public int pitch { get; set; }
    public int volume { get; set; }
  }

  public class Airship
  {
    public Sound bgm { get; set; }
    public int characterIndex { get; set; }
    public string characterName { get; set; }
    public int startMapId { get; set; }
    public int startX { get; set; }
    public int startY { get; set; }
  }

  public class AttackMotion
  {
    public int type { get; set; }
    public int weaponImageId { get; set; }
  }

  public class Boat
  {
    public Sound bgm { get; set; }
    public int characterIndex { get; set; }
    public string characterName { get; set; }
    public int startMapId { get; set; }
    public int startX { get; set; }
    public int startY { get; set; }
  }

  public class Ship
  {
    public Sound bgm { get; set; }
    public int characterIndex { get; set; }
    public string characterName { get; set; }
    public int startMapId { get; set; }
    public int startX { get; set; }
    public int startY { get; set; }
  }

  public class Messages
  {
    public string actionFailure { get; set; }
    public string actorDamage { get; set; }
    public string actorDrain { get; set; }
    public string actorGain { get; set; }
    public string actorLoss { get; set; }
    public string actorNoDamage { get; set; }
    public string actorNoHit { get; set; }
    public string actorRecovery { get; set; }
    public string alwaysDash { get; set; }
    public string bgmVolume { get; set; }
    public string bgsVolume { get; set; }
    public string buffAdd { get; set; }
    public string buffRemove { get; set; }
    public string commandRemember { get; set; }
    public string counterAttack { get; set; }
    public string criticalToActor { get; set; }
    public string criticalToEnemy { get; set; }
    public string debuffAdd { get; set; }
    public string defeat { get; set; }
    public string emerge { get; set; }
    public string enemyDamage { get; set; }
    public string enemyDrain { get; set; }
    public string enemyGain { get; set; }
    public string enemyLoss { get; set; }
    public string enemyNoDamage { get; set; }
    public string enemyNoHit { get; set; }
    public string enemyRecovery { get; set; }
    public string escapeFailure { get; set; }
    public string escapeStart { get; set; }
    public string evasion { get; set; }
    public string expNext { get; set; }
    public string expTotal { get; set; }
    public string file { get; set; }
    public string levelUp { get; set; }
    public string loadMessage { get; set; }
    public string magicEvasion { get; set; }
    public string magicReflection { get; set; }
    public string meVolume { get; set; }
    public string obtainExp { get; set; }
    public string obtainGold { get; set; }
    public string obtainItem { get; set; }
    public string obtainSkill { get; set; }
    public string partyName { get; set; }
    public string possession { get; set; }
    public string preemptive { get; set; }
    public string saveMessage { get; set; }
    public string seVolume { get; set; }
    public string substitute { get; set; }
    public string surprise { get; set; }
    public string useItem { get; set; }
    public string victory { get; set; }
  }

  public class Terms
  {
    public List<string> basic { get; set; }
    public List<string> commands { get; set; }
    public List<string> @params { get; set; }
    public Messages messages { get; set; }
  }

  public class TestBattler
  {
    public int actorId { get; set; }
    public List<int?> equips { get; set; }
    public int level { get; set; }
  }

  public class Timing
  {
    public List<int> flashColor { get; set; }
    public int flashDuration { get; set; }
    public int flashScope { get; set; }
    public int frame { get; set; }
    public Sound se { get; set; }
  }

  public class ActionList
  {
    public int code { get; set; }
    public int indent { get; set; }
    public List<object> parameters { get; set; }
  }

  public class RMAction
  {
    public double conditionParam1 { get; set; }
    public double conditionParam2 { get; set; }
    public int conditionType { get; set; }
    public int rating { get; set; }
    public int skillId { get; set; }
  }

  public class DropItem
  {
    /// <summary>
    /// 0: 无, 1: 物品, 2: 武器, 3: 护甲
    /// </summary>
    public int kind { get; set; }

    /// <summary>
    /// 对应 物品/武器/护甲 的数据ID
    /// </summary>
    public int dataId { get; set; }

    /// <summary>
    /// 几率 = 1 / {denominator}
    /// </summary>
    public int denominator { get; set; }

    public static DropItem Empty => new DropItem {kind = 0, dataId = 0, denominator = 0};
  }

  public class Goods
  {
    public List<int> items { get; set; }
    public List<int> weapons { get; set; }
    public List<int> armors { get; set; }
  }
}
