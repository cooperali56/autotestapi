from abc import ABC, abstractmethod


class ApiSetData(ABC):
    """
    接口前数据处理
    """

    @abstractmethod
    def api_data_to_dict(self, *args):
        """
        参数转字典

        Args:
            args: key集

        Returns:
            None
        """
        pass
