from sqlite3 import connect


class CollegeNode:
    """
    院校节点

    :param ID: 代号
    :param Name: 院校名称
    :param Number: XX年计划数 [2021, 2022, ... ]
    """

    def __init__(self,
                 ID: str,
                 Name: str,
                 Number: list):
        self.ID = ID
        self.Name = Name
        self.Number = Number


class GroupNode:
    """
    专业组节点

    :param ID: 代号
    :param Parent: 所属院校
    :param Name: 专业组名称
    :param Limitation: 限制
    :param Number: XX年计划数[2021, 2022, ... ]
    :param Line: XX年分数线[2021, 2022, ... ]
    """

    def __init__(self,
                 ID: str,
                 Parent: str,
                 Name: str,
                 Limitation: str,
                 Number: list,
                 Line: list):
        self.ID = ID
        self.Parent = Parent
        self.Name = Name
        self.Limitation = Limitation
        self.Number = Number
        self.Line = Line


class SpecializationNode:
    """
    专业节点

    :param ID: 代号
    :param Parent: 所属专业组
    :param Name: 专业名称
    :param Time: 学制
    :param Fee: 学费
    :param Number: XX年计划数[2021, 2022, ... ]
    """

    def __init__(self,
                 ID: str,
                 Parent: str,
                 Name: str,
                 Time: str,
                 Fee: str,
                 Number: list):
        self.ID = ID
        self.Parent = Parent
        self.Name = Name
        self.Time = Time
        self.Fee = Fee
        self.Number = Number


class ParameterInvalid(Exception):
    def __init__(self, User: str, Demand: str):
        self.User = User
        self.Demand = Demand

    def __str__(self):
        return '参数无效，应为 {} ，输入 {} '.format(self.Demand, self.User)


class UnavailableID(Exception):
    def __init__(self, User: str):
        self.User = User

    def __str__(self):
        return '无法找到院校代号为 {} 的记录'.format(self.User)


class DataTree:
    """
    构造树

    :param ID: 院校代号
    :param Subject: 科目 ( wenke 或 like )
    :param Type: 类型 ( global 或 local )

    :except ParameterInvalid: 参数 Subject 或 Type 不符合规则
    :except UnavailableID: 无法在数据库中找到对应院校编号的记录
    """

    def __init__(self, ID: str, Subject: str, Type: str):
        self.ID = ID
        self.Subject = Subject
        self.Type = Type
        self.Tree = {}
        """
        Tree: {
          CollegeNode: {
            GroupNode1: [
              SpecializationNode1,
              SpecializationNode2,
              SpecializationNode3,
                    ... ...
              ],
            GroupNode2: [
                    ... ...
            ],
            GroupNode3: [
                    ... ...
            ], 
                ... ...
            }
        }
        """

        if self.Subject not in ['wenke', 'like']:
            raise ParameterInvalid(self.Subject, 'wenke 或 like')
        if self.Type not in ['global', 'local']:
            raise ParameterInvalid(self.Type, 'global 或 local')

        self._BuildCollegeNodes()

    def _BuildCollegeNodes(self):
        GaokaoDB = connect('gaokaoDB')
        Datas = GaokaoDB.execute(
            'select * from {}_{} where 代号="{}"'
            .format(self.Subject, self.Type, self.ID)
        )
        for Data in Datas:
            Data = list(Data)
            Node = CollegeNode(Data.pop(0), Data.pop(0), Data)
            self.Tree[Node] = {}
            self._BuildGroupNodes(self.ID, Node)
        GaokaoDB.close()

        if self.Tree == {}:
            raise UnavailableID(self.ID)

    def _BuildGroupNodes(self, Parent, NodeParent):
        GaokaoDB = connect('gaokaoDB')
        Datas = GaokaoDB.execute(
            'select * from {}_{}_spGroups where 所属院校="{}" order by 代号'
            .format(self.Subject, self.Type, Parent)
        )
        for Data in Datas:
            Data = list(Data)
            ID = Data.pop(0)
            Node = GroupNode(ID, Data.pop(0), Data.pop(0), Data.pop(0), Data[: int(len(Data) / 2)],
                             Data[int(len(Data) / 2):])
            self.Tree[NodeParent][Node] = []
            self._BuildSpecializationNodes(ID, Node, NodeParent)
        GaokaoDB.close()

    def _BuildSpecializationNodes(self, Parent, NodeParent, NodeAncestor):
        GaokaoDB = connect('gaokaoDB')
        Datas = GaokaoDB.execute(
            'select * from {}_{}_sps where 所属专业组="{}" order by 代号'
            .format(self.Subject, self.Type, Parent)
        )
        for Data in Datas:
            Data = list(Data)
            Data.pop(2)
            Node = SpecializationNode(Data.pop(0), Data.pop(0), Data.pop(0), Data.pop(0), Data.pop(0), Data)
            self.Tree[NodeAncestor][NodeParent].append(Node)
        GaokaoDB.close()

    def GetCollegeNode(self) -> CollegeNode:
        """
        获取院校节点

        :return: 院校节点
        """
        for Root in self.Tree:
            return Root

    def GetGroupNodes(self) -> list[GroupNode]:
        """
        获取所有专业组节点

        :return: 包含所有专业组节点的列表
        """
        Result = []
        for Root in self.Tree:
            for Group in self.Tree[Root]:
                Result.append(Group)
        return Result

    def GetSpecializationNodes(self, IndexOn: bool = True) \
            -> dict[str, list[SpecializationNode]] | list[SpecializationNode]:
        """
        获取所有专业节点

        :param IndexOn: 是否为专业节点附加所属专业组代号作为索引，默认为 True
        :return: 若IndexOn为真，则返回形如{ 专业组代号1: [专业1, 专业2, ... ...], ... ... }的字典，否则返回由来自不同专业组的所有专业构成的无序列表
        """
        if IndexOn:
            Result = {}
            for Root in self.Tree:
                for Group in self.Tree[Root]:
                    Result[Group.ID] = self.Tree[Root][Group]
            return Result
        else:
            Result = []
            for Root in self.Tree:
                for Group in self.Tree[Root]:
                    Result += self.Tree[Root][Group]
            return Result
